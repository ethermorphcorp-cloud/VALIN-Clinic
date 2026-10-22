/* =====================================================================
 * VALIN CLINIC — Google Apps Script backend
 * =====================================================================
 * Files in this project:
 *   Code.gs       — this file. All server-side logic.
 *   index.html    — the main single-page app (client). Self-contained
 *                   HTML/CSS/JS, calls server functions via
 *                   google.script.run.
 *   Receipt.html  — printable receipt page, served via
 *                   doGet(?page=receipt&ticket=...).
 *   Line.html     — HTML template for the LINE daily-summary message
 *                   body (evaluated server-side, sent as text via the
 *                   LINE Messaging API — LINE Notify was shut down by
 *                   LINE on 2025-03-31 and can no longer be used).
 *
 * DATA STORE
 * ----------
 * A single Google Spreadsheet (created automatically by setupProject()
 * the first time it runs) holds one sheet/tab per entity. This list is
 * kept in sync with the literal source of truth, `CONFIG.HEADERS` below —
 * (round 20 cleanup: this comment had drifted badly out of date across 16
 * rounds of append-only schema growth, e.g. it was still missing every
 * Records column added since round 4, and the entire Promotions sheet;
 * if the two ever disagree again, `CONFIG.HEADERS` is what's actually
 * running — treat this comment as documentation, not configuration):
 *
 *   Patients       ID, FirstName, LastName, Phone, DOB, Gender, Email,
 *                  PhotoFileId, JoinedDate, CreatedAt, UpdatedAt, Nickname
 *   Records        Core: RID, Date, PatientID, PatientName, Phone,
 *                  ServiceCategory, Detail, DoctorID, DoctorName, SalesID,
 *                  SalesName, PaymentMethod, SuppliesJSON, UseCanula,
 *                  CanulaAmount, Amount, Status, ReceiptNo, Note,
 *                  CreatedBy, CreatedAt, UpdatedAt.
 *                  Revenue-split (round 4/13, see "Revenue-split fields"
 *                  comment on CONFIG.HEADERS.Records below for the full
 *                  formulas): OtherAmount, SalesCommissionPct, SalesShare,
 *                  DoctorCommissionPct, DoctorShareRaw, DoctorShare,
 *                  DoctorShareNote, ClinicShare, VisitStatus.
 *                  Round 13 itemized lines (each *JSON is an array, see
 *                  parseLineItems_): UsePromo, PromoJSON, PromoAmount,
 *                  UseMedicine, MedicineJSON, MedicineAmount, UseOther,
 *                  OtherItemsJSON, SupplierPayAmount.
 *                  Round 14: OpdImageFileId (optional Drive attachment,
 *                  no revenue-split involvement at all).
 *   Stock          ItemID, Name, Category, Unit, Qty, CreatedAt, UpdatedAt —
 *                  Qty (round 25) is a denormalized cache of the touched
 *                  item's TRUE latest StockLog entry by date, kept in sync
 *                  by resyncStockQtyFromLog_() on every StockLog write.
 *   StockLog       (round 25) LogID, Date, ItemID, ItemName, QtyIn,
 *                  QtyRemaining, Note, CreatedBy, CreatedAt — one row per
 *                  item per "อัพเดตสต๊อก" batch save (Date/Note repeated
 *                  across every item line in that same batch, no separate
 *                  header sheet — see saveStockUpdate below). Feeds the
 *                  Stock page's per-item history sub-table and its
 *                  "เรียงตาม update ล่าสุด" sort.
 *   Expenses       ExpID, Date, Name, Category, Amount, CreatedBy, CreatedAt
 *   Orders         PoID, Date, Supplier, TotalAmount, CreatedBy, CreatedAt
 *   OrderLines     LineID, PoID, Category, ItemName, Qty, UnitPrice, LineTotal
 *   Doctors        DoctorID, Name, Status, MinAmount, CreatedAt, CommissionPct
 *   Sales          SalesID, Name, Status, CommissionRate, MinAmount, CreatedAt
 *   Promotions     (round 12/13) PromoID, Name, Price, DoctorSharePct,
 *                  Status, CreatedAt, UpdatedAt — feeds the record modal's
 *                  "รายการโปร" picker and the Records.PromoJSON lines above;
 *                  ClinicSharePct is always derived as (100 - DoctorSharePct)
 *                  at read time, never stored.
 *   Categories     CategoryID, Type, Name, SortOrder      (Type: treatment|stock|expense)
 *   Staff          StaffID, Name, Role, Username, PasswordHash, PasswordSalt,
 *                  LineID, Level, Active, CreatedAt — LineID (round 20) is
 *                  what the LINE webhook chat commands authorize against,
 *                  see doPost()/findLineAuthorizedStaff_() below. Round 21:
 *                  LineID + Level=admin + Active together are also now what
 *                  the "ส่งข้อมูลไป LINE" push button sends to, see
 *                  getLineAdminTargets_() below. Role is a separate,
 *                  free-text job-title field (e.g. "ผู้ดูแลระบบ") — NOT the
 *                  same thing as Level, which is the actual admin/staff
 *                  auth level checked everywhere else in this file.
 *   ClinicSettings Key, Value    (single key/value table: name, phone, address, ...)
 *   Counters       Key, LastSeq  (used for receipt-number sequencing)
 *   AuditLog       Timestamp, User, Action, Entity, EntityID, Detail
 *
 * NOTE ON ORDERS: a purchase order (PO) has one header row (date,
 * supplier, grand total) in `Orders`, and one row per line item
 * (category, item, qty, unit price, line total) in `OrderLines`,
 * linked by PoID. The UI table lists one row per PO; editing a PO
 * opens all of its lines together, matching the "รองรับมากกว่า 1
 * รายการ" requirement in the order modal spec.
 *
 * SETUP (run once from the Apps Script editor):
 *   1. Run `setupProject` — creates the spreadsheet, all sheets with
 *      headers, the Drive folder structure, and default Script
 *      Properties. Authorize the requested scopes when prompted.
 *   2. (Optional) Run `seedDemoData` to fill the sheets with the same
 *      sample data used in the front-end prototype.
 *   3. Deploy > New deployment > Web app. Execute as "Me", access
 *      "Anyone" (or "Anyone within your organization").
 *   4. Set Script Properties (Project Settings > Script Properties):
 *        LINE_CHANNEL_ACCESS_TOKEN  — from your LINE Official Account
 *      (Round 21: the old LINE_TARGET_ID property is no longer used — see
 *      step 4a.)
 *   4a. (Round 21) The "ส่งข้อมูลไป LINE" push button no longer sends to one
 *      fixed saved id. It now sends to EVERY Staff record whose Level is
 *      "admin" AND whose "LINE ID" field (ตั้งค่า > ตั้งค่าพนักงาน) is
 *      filled in and who is Active — have each admin message the bot
 *      "Line ID ของฉัน" (see step 5 below) to learn their own LINE user id,
 *      then enter it into their own staff record. Until at least one admin
 *      has done this, the push button will fail with a clear error instead
 *      of silently doing nothing.
 *   5. (Round 20 item 2, optional — only needed for the interactive LINE
 *      chat commands, not the "ส่งข้อมูลไป LINE" push button, which only
 *      needs step 4/4a) Set the Webhook URL in the LINE Developers Console,
 *      under your channel's Messaging API tab, and turn "Use webhook" ON.
 *      **(Round 23) Use a SEPARATE deployment's URL here, not the same
 *      /exec URL staff use to log into the web app**: from "Deploy" >
 *      "Manage deployments" > "New deployment" (Web app, Execute as "Me",
 *      access "Anyone"), create a second deployment and copy ITS /exec URL
 *      for the LINE Webhook URL field — keep the original deployment's URL
 *      (step 3) as what staff use. Both deployments run the exact same
 *      `doGet`/`doPost` code (nothing is actually split functionally —
 *      GAS has no way to give one script two different behaviors per URL),
 *      this purely gives LINE its own address separate from the one you
 *      hand out to staff. **Whenever you push a code update, re-select the
 *      latest version on BOTH deployments** (Manage deployments > pencil
 *      icon on each > "Version: New version" > Deploy) — updating only one
 *      leaves the other silently running stale code. Then, for each staff
 *      member who should be allowed to use the chat commands, have them
 *      message the bot "Line ID ของฉัน" to learn their LINE user id, and
 *      enter it into their staff record's "LINE ID" field (ตั้งค่า >
 *      ตั้งค่าพนักงาน) — see doPost()/handleLineWebhookEvent_() below for
 *      the full command list and authorization rule. **Note**: the
 *      chat-bot commands remain open to any Active staff member (not just
 *      admins) per Round 20's design (assumption #31 in the project
 *      notes) — only the PUSH button's recipient list was scoped to
 *      admins in Round 22.
 * =====================================================================
 */

/* ---------------------------------------------------------------- */
/* CONFIG                                                             */
/* ---------------------------------------------------------------- */
var CONFIG = {
  SHEETS: {
    PATIENTS: 'Patients',
    RECORDS: 'Records',
    STOCK: 'Stock',
    STOCK_LOG: 'StockLog',
    EXPENSES: 'Expenses',
    ORDERS: 'Orders',
    ORDER_LINES: 'OrderLines',
    DOCTORS: 'Doctors',
    SALES: 'Sales',
    PROMOTIONS: 'Promotions',
    CATEGORIES: 'Categories',
    STAFF: 'Staff',
    CLINIC_SETTINGS: 'ClinicSettings',
    COUNTERS: 'Counters',
    MED_DEPOSITS: 'MedDeposits',
    AUDIT_LOG: 'AuditLog'
  },
  HEADERS: {
    // NOTE ON SCHEMA CHANGES: new columns are always appended at the END of
    // these arrays, never inserted/renamed in the middle — ensureSheet_()
    // rewrites row 1 (headers only, never data) whenever this array doesn't
    // match what's on the sheet, so appending keeps every existing row's
    // values correctly aligned under their original columns. Re-running
    // setupProject() after a schema change like this is always safe.
    Patients: ['ID', 'FirstName', 'LastName', 'Phone', 'DOB', 'Gender', 'Email', 'PhotoFileId', 'JoinedDate', 'CreatedAt', 'UpdatedAt', 'Nickname'],
    // Revenue-split fields (added for the commission engine): OtherAmount is
    // now the SUM of the round-13 repeatable "รายการอื่นๆ" lines (see
    // OtherItemsJSON) — still excluded from anyone's commission share, only
    // feeds SupplierPayAmount. SalesCommissionPct/DoctorCommissionPct
    // snapshot the rate actually used at save time (so later rate changes on
    // the Doctors/Sales sheets don't retroactively rewrite historical
    // splits). DoctorShareRaw is the plain commission BEFORE the
    // daily-minimum top-up — as of round 13 it is
    // Amount*DoctorCommissionPct/100 + sum(each PromoJSON line's amount *
    // that promo's own DoctorSharePct/100) (see computeRecordShares_).
    // DoctorShare is the FINAL amount actually used (topped up if needed);
    // DoctorShareNote records when/why a top-up was applied. ClinicShare =
    // (Amount + PromoAmount + CanulaAmount) - SalesShare - DoctorShare —
    // i.e. whatever's left of the procedure amount + promo amount after the
    // sales/doctor cuts, plus canula in full (see recomputeDoctorDayShares_).
    // VisitStatus is a separate นัดหมาย/เสร็จสิ้น workflow marker, independent
    // of the payment Status column.
    //
    // Round 13 additions (appended, never reorder/rename existing columns):
    // UsePromo/PromoJSON/PromoAmount — checkbox-gated, repeatable "รายการโปร
    // หน้าคลีนิค" lines, each {promoId,name,price,doctorSharePct,amount};
    // PromoAmount is the denormalized sum of each line's `amount`, used
    // directly in the share formulas above and by recomputeDoctorDayShares_.
    // UseMedicine/MedicineJSON/MedicineAmount — checkbox-gated, repeatable
    // free-text "ซื้อเพิ่มสินค้า/เบิกยาหน้าคลีนิค" lines, each {name,amount} —
    // deliberately separate from the existing stock-linked SuppliesJSON
    // picker (which still deducts Stock qty on save); these medicine lines
    // are NOT linked to inventory, they only feed SupplierPayAmount.
    // UseOther/OtherItemsJSON — the itemized version of "รายการอื่นๆ", each
    // {label,amount}; OtherAmount (existing column, above) is kept as the
    // denormalized sum for backward compatibility.
    // SupplierPayAmount — denormalized MedicineAmount + OtherAmount ("จ่าย
    // supplier" in the record modal); a DIFFERENT concept from the
    // dashboard's Orders-based "ยอดจ่าย Supplier" slice (purchase orders) —
    // the two are not reconciled/merged.
    Records: ['RID', 'Date', 'PatientID', 'PatientName', 'Phone', 'ServiceCategory', 'Detail', 'DoctorID', 'DoctorName', 'SalesID', 'SalesName', 'PaymentMethod', 'SuppliesJSON', 'UseCanula', 'CanulaAmount', 'Amount', 'Status', 'ReceiptNo', 'Note', 'CreatedBy', 'CreatedAt', 'UpdatedAt', 'OtherAmount', 'SalesCommissionPct', 'SalesShare', 'DoctorCommissionPct', 'DoctorShareRaw', 'DoctorShare', 'DoctorShareNote', 'ClinicShare', 'VisitStatus', 'UsePromo', 'PromoJSON', 'PromoAmount', 'UseMedicine', 'MedicineJSON', 'MedicineAmount', 'UseOther', 'OtherItemsJSON', 'SupplierPayAmount', 'OpdImageFileId', 'UseSupplies', 'UseOwnMed', 'OwnMedJSON', 'UseDepositUse', 'DepositUseJSON', 'UseMessenger', 'MessengerJSON', 'StockApplied'],
    // MedDeposits — ยาฝาก: one balance row per patient + medicine item, kept
    // in sync by applyRecordEffects_ (เบิกยาหน้าคลีนิค lines add their จำนวนฝาก,
    // รายการใช้ยาฝาก lines take it back out). Date/CreatedBy are the latest change.
    MedDeposits: ['DepositID', 'Date', 'PatientID', 'PatientName', 'ItemName', 'QtyRemaining', 'CreatedBy', 'ItemID', 'UpdatedAt'],
    // StockType: 'supply' (เวชภัณฑ์) | 'medicine' (ยา) — splits the Stock page
    // into two tabs and feeds the record modal's stock-backed name
    // suggestions (รายการเวชภัณฑ์ที่ใช้ / รายการซื้อยาเพิ่ม). Blank on rows
    // created before the split, which read as 'supply' (see stockTypeOf_).
    // MinQty = จำนวนสต๊อกขั้นต่ำ (drives the ใกล้หมด badge); CostPrice =
    // ราคาทุน (admin-only — stripped from every non-admin response and never
    // overwritten by a non-admin save); Price1..Price50 = ราคาขาย per quantity
    // tier, see STOCK_PRICE_TIERS.
    Stock: ['ItemID', 'Name', 'Category', 'Unit', 'Qty', 'CreatedAt', 'UpdatedAt', 'StockType', 'MinQty', 'CostPrice', 'Price1', 'Price3', 'Price5', 'Price10', 'Price20', 'Price50'],
    // StockLog (round 25) — one row per item PER "อัพเดตสต๊อก" batch save; a
    // single save (one Date + one Note) can spawn several StockLog rows,
    // one per item line, sharing that Date/Note (mirrors the
    // Orders/OrderLines header+lines split, except here there's no separate
    // header sheet — Date/Note are simply repeated on each line row, since
    // unlike a PO there's no other header-only field like Supplier/Total to
    // justify a second sheet). QtyIn = จำนวนรับเข้า for this entry;
    // QtyRemaining = จำนวนคงเหลือ the staff counted/typed for this entry
    // (NOT auto-computed as a running total — round 25's spec has the user
    // type the remaining count directly each time, e.g. after a physical
    // stock count). Stock.Qty (above) is kept in sync as a denormalized
    // "latest QtyRemaining across this item's StockLog rows" on every save
    // (see saveStockUpdate below), so listStock's existing sort/filter/
    // display logic and any other reader of Stock.Qty keeps working
    // unchanged; the Stock page's new table always shows the latest
    // StockLog entry's QtyRemaining per item, sorted by that entry's Date.
    // QtyOut/Kind appended later: Kind is what produced the row (see
    // STOCK_LOG_KINDS); rows written before QtyOut existed have it blank and
    // getStockLog derives it from the previous row's QtyRemaining instead.
    StockLog: ['LogID', 'Date', 'ItemID', 'ItemName', 'QtyIn', 'QtyRemaining', 'Note', 'CreatedBy', 'CreatedAt', 'QtyOut', 'Kind'],
    Expenses: ['ExpID', 'Date', 'Name', 'Category', 'Amount', 'CreatedBy', 'CreatedAt'],
    Orders: ['PoID', 'Date', 'Supplier', 'TotalAmount', 'CreatedBy', 'CreatedAt'],
    OrderLines: ['LineID', 'PoID', 'Category', 'ItemName', 'Qty', 'UnitPrice', 'LineTotal'],
    Doctors: ['DoctorID', 'Name', 'Status', 'MinAmount', 'CreatedAt', 'CommissionPct'],
    Sales: ['SalesID', 'Name', 'Status', 'CommissionRate', 'MinAmount', 'CreatedAt'],
    // Promotions (round 12, item 3) — reference/display table for the
    // "จัดการโปรโมชันหน้าคลินิก" page. NOTE: round 12 deliberately kept this
    // standalone from the treatment record modal / revenue-split engine; round
    // 13 explicitly REVERSES that — the record modal's "รายการโปร" picker now
    // reads this list live, and each promo's DoctorSharePct feeds directly
    // into ส่วนแบ่งแพทย์ (see computeRecordShares_ / Records.PromoJSON above).
    // ClinicSharePct is always derived as (100 - DoctorSharePct), never
    // stored, so it can never drift.
    Promotions: ['PromoID', 'Name', 'Price', 'DoctorSharePct', 'Status', 'CreatedAt', 'UpdatedAt'],
    Categories: ['CategoryID', 'Type', 'Name', 'SortOrder'],
    Staff: ['StaffID', 'Name', 'Role', 'Username', 'PasswordHash', 'PasswordSalt', 'LineID', 'Level', 'Active', 'CreatedAt'],
    ClinicSettings: ['Key', 'Value'],
    Counters: ['Key', 'LastSeq'],
    AuditLog: ['Timestamp', 'User', 'Action', 'Entity', 'EntityID', 'Detail']
  },
  // Columns that must be stored/read as literal text so leading zeros
  // (phone numbers, ID codes) are never lost to number coercion.
  TEXT_COLUMNS: {
    Patients: ['Phone'],
    Records: ['Phone', 'ReceiptNo'],
    Staff: ['Username']
  },
  PROPS: {
    SPREADSHEET_ID: 'SPREADSHEET_ID',
    PHOTOS_FOLDER_ID: 'PHOTOS_FOLDER_ID',
    RECEIPTS_FOLDER_ID: 'RECEIPTS_FOLDER_ID',
    // Round 17 item 2: OPD photos now live in their own Year/Month/Day tree
    // (like Receipts already did), separate from the flat "Photos" folder
    // (patient photos / login image / clinic logo) so the daily structure
    // doesn't mix with those un-dated, flat-named files.
    OPD_PHOTOS_FOLDER_ID: 'OPD_PHOTOS_FOLDER_ID',
    ROOT_FOLDER_ID: 'ROOT_FOLDER_ID',
    // Round 21: LINE_TARGET_ID (a single fixed push target) was removed —
    // the "ส่งข้อมูลไป LINE" push now multicasts to every admin-level Staff
    // row that has its own LineID filled in, via getLineAdminTargets_()
    // below, instead of one shared Script Property. Only the channel
    // access token is still needed as a Script Property.
    LINE_CHANNEL_ACCESS_TOKEN: 'LINE_CHANNEL_ACCESS_TOKEN'
  },
  CACHE: {
    MASTER_DATA_KEY: 'MASTER_DATA_V1',
    MASTER_DATA_TTL: 21600, // 6h, the max CacheService allows
    SESSION_PREFIX: 'SESSION_',
    SESSION_TTL: 21600,     // 6h
    RECEIPT_TICKET_PREFIX: 'RTKT_',
    RECEIPT_TICKET_TTL: 300, // 5 minutes
    // Lightweight full-list caches used to power client-side autocomplete
    // (patient/stock suggest boxes) without a server round-trip per keystroke.
    PATIENTS_LITE_KEY: 'PATIENTS_LITE_V1',
    STOCK_LITE_KEY: 'STOCK_LITE_V3', // V3: entries carry `type` and sale `prices` (never cost)
    LITE_TTL: 21600,
    // Round 19 item 4: getDashboardData's per-date-range result cache. Keyed
    // by a version tag (DASH_VERSION_KEY, bumped by any write that could
    // change dashboard numbers — see bumpDashboardCacheVersion_) rather than
    // invalidated by key removal, since the actual cache key also varies by
    // date range and CacheService has no prefix/wildcard delete — bumping
    // the version makes every previously-cached range unreachable at once.
    DASH_KEY_PREFIX: 'DASH_V1_',
    DASH_VERSION_KEY: 'DASH_VERSION_V1',
    DASH_TTL: 1800 // 30 min — safe because correctness comes from the version bump, not the TTL
  },
  PAGE_SIZE_DEFAULT: 10
};

/* ---------------------------------------------------------------- */
/* Web app entry point                                                */
/* ---------------------------------------------------------------- */
function doGet(e) {
  var page = (e && e.parameter && e.parameter.page) || 'app';
  if (page === 'receipt') {
    return renderReceiptPage_(e.parameter.ticket);
  }
  var tpl = HtmlService.createTemplateFromFile('index');
  tpl.loginImageUrl = getLoginImageUrl_();
  return tpl.evaluate()
    .setTitle('VALIN CLINIC | ระบบบริหารคลินิก')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function renderReceiptPage_(ticket) {
  var cache = CacheService.getScriptCache();
  var raw = ticket ? cache.get(CONFIG.CACHE.RECEIPT_TICKET_PREFIX + ticket) : null;
  var tpl = HtmlService.createTemplateFromFile('Receipt');
  if (!raw) {
    tpl.valid = false;
    tpl.data = null;
  } else {
    var rid = JSON.parse(raw).rid;
    tpl.valid = true;
    tpl.data = buildReceiptData_(rid);
  }
  return tpl.evaluate().setTitle('ใบเสร็จรับเงิน | VALIN CLINIC');
}

/* ---------------------------------------------------------------- */
/* Spreadsheet / sheet helpers                                        */
/* ---------------------------------------------------------------- */
// Round 17 item 4 (performance): every server function that touches more
// than one sheet — which is most of them (saveRecord alone looks up
// RECORDS, then STOCK, then re-reads RECORDS to recompute doctor shares,
// etc.) — used to call SpreadsheetApp.openById() and getScriptProperties()
// fresh on every single getSheet_()/getSS_() call. Each Apps Script web-app
// request runs in its own fresh execution (top-level vars never carry
// data between separate requests), so memoizing the opened Spreadsheet —
// and each looked-up Sheet — for the lifetime of THIS request only cuts
// pure repeat-lookup overhead; it can never serve stale data across
// requests, and within a request the underlying sheet identity never
// changes mid-call anyway (only its cell VALUES do, and those are always
// re-read fresh via getDataRange()/getRows_ regardless of this cache).
var _ssCache_ = null;
function getSS_() {
  if (_ssCache_) return _ssCache_;
  var props = PropertiesService.getScriptProperties();
  var id = props.getProperty(CONFIG.PROPS.SPREADSHEET_ID);
  if (!id) throw new Error('ยังไม่ได้ตั้งค่าโปรเจกต์ กรุณารัน setupProject() ก่อน (SPREADSHEET_ID ไม่พบใน Script Properties)');
  _ssCache_ = SpreadsheetApp.openById(id);
  return _ssCache_;
}

function ensureSheet_(ss, name, headers) {
  var sh = ss.getSheetByName(name);
  if (!sh) {
    sh = ss.insertSheet(name);
  }
  var firstRow = sh.getRange(1, 1, 1, headers.length).getValues()[0];
  var needsHeader = headers.some(function (h, i) { return firstRow[i] !== h; });
  if (needsHeader) {
    sh.getRange(1, 1, 1, headers.length).setValues([headers]);
    sh.setFrozenRows(1);
    sh.getRange(1, 1, 1, headers.length).setFontWeight('bold').setBackground('#E9EDF7');
  }
  // Force text-only columns to plain-text number format so values like
  // "081-234-5678" or "0812345678" never lose a leading zero.
  var textCols = CONFIG.TEXT_COLUMNS[name] || [];
  textCols.forEach(function (col) {
    var idx = headers.indexOf(col);
    if (idx > -1) {
      sh.getRange(1, idx + 1, Math.max(sh.getMaxRows(), 1000), 1).setNumberFormat('@');
    }
  });
  return sh;
}

// Same per-request memoization as getSS_ above, keyed by sheet name — the
// Sheet *object reference* is cheap to reuse; cell data is still always
// read fresh (getRows_ calls getDataRange().getValues() every time it's
// invoked, cache or no cache), so this cannot serve stale rows.
var _sheetCache_ = {};
function getSheet_(name) {
  if (_sheetCache_[name]) return _sheetCache_[name];
  var ss = getSS_();
  var sh = ss.getSheetByName(name);
  // A sheet added to CONFIG after setup (e.g. MedDeposits) is created on first use.
  if (!sh && CONFIG.HEADERS[name]) sh = ensureSheet_(ss, name, CONFIG.HEADERS[name]);
  if (!sh) throw new Error('ไม่พบชีต: ' + name + ' — กรุณารัน setupProject() ก่อน');
  // Columns are only ever appended (see CONFIG.HEADERS), so a header row
  // narrower than the schema means a newly-appended column hasn't been
  // written to row 1 yet — write it now instead of requiring a manual
  // setupProject() re-run after every schema change.
  var expected = CONFIG.HEADERS[name];
  if (expected && sh.getLastColumn() < expected.length) ensureSheet_(ss, name, expected);
  _sheetCache_[name] = sh;
  return sh;
}

// Reads all data rows as an array of plain objects keyed by header name.
// Skips rows whose first column is blank (avoids GAS's "phantom" empty
// rows at the bottom of getDataRange()).
function getRows_(sheet) {
  var values = sheet.getDataRange().getValues();
  if (values.length < 2) return [];
  var headers = values[0];
  var rows = [];
  for (var r = 1; r < values.length; r++) {
    var row = values[r];
    if (row[0] === '' || row[0] === null) continue;
    var obj = {};
    for (var c = 0; c < headers.length; c++) obj[headers[c]] = row[c];
    obj._row = r + 1; // 1-based sheet row number, for update/delete
    rows.push(obj);
  }
  return rows;
}

function findRowIndexById_(sheet, idCol, id) {
  var rows = getRows_(sheet);
  for (var i = 0; i < rows.length; i++) {
    if (String(rows[i][idCol]) === String(id)) return rows[i]._row;
  }
  return -1;
}

/**
 * Applies the '@' (plain text) number format to a row's TEXT_COLUMNS cells
 * BEFORE the values are written. ensureSheet_() only pre-formats the first
 * ~1000 rows at setup time, so relying on that alone silently breaks once a
 * sheet grows past it — Sheets auto-detects a numeric-looking string
 * ("0812345678") on a still-"Automatic"-formatted cell and strips the
 * leading zero. Setting the format immediately before every single write
 * (append or update) makes this correct regardless of sheet size.
 */
function applyTextColumnFormat_(sheet, headers, rowIndex) {
  var textCols = CONFIG.TEXT_COLUMNS[sheet.getName()] || [];
  textCols.forEach(function (col) {
    var idx = headers.indexOf(col);
    if (idx > -1) sheet.getRange(rowIndex, idx + 1, 1, 1).setNumberFormat('@');
  });
}

function appendRow_(sheet, headers, obj) {
  var rowIndex = sheet.getLastRow() + 1;
  applyTextColumnFormat_(sheet, headers, rowIndex);
  var row = headers.map(function (h) { return obj[h] !== undefined ? obj[h] : ''; });
  sheet.getRange(rowIndex, 1, 1, headers.length).setValues([row]);
  return rowIndex;
}

function updateRowByIndex_(sheet, headers, rowIndex, obj) {
  applyTextColumnFormat_(sheet, headers, rowIndex);
  var row = headers.map(function (h) { return obj[h] !== undefined ? obj[h] : ''; });
  sheet.getRange(rowIndex, 1, 1, headers.length).setValues([row]);
}

function deleteRowByIndex_(sheet, rowIndex) {
  sheet.deleteRow(rowIndex);
}

/* Generic paginate + simple text search over an already-filtered array */
function paginate_(items, page, pageSize) {
  pageSize = pageSize || CONFIG.PAGE_SIZE_DEFAULT;
  page = Math.max(1, page || 1);
  var total = items.length;
  var pages = Math.max(1, Math.ceil(total / pageSize));
  page = Math.min(page, pages);
  var start = (page - 1) * pageSize;
  return { items: items.slice(start, start + pageSize), total: total, page: page, pages: pages, pageSize: pageSize };
}

/* ======================================================================
   SETUP — run setupProject() once manually from the Apps Script editor
   ====================================================================== */

/**
 * Idempotent project bootstrap.
 * - Creates a bound/target Spreadsheet if SPREADSHEET_ID script property is not set
 *   (or uses the one already configured), then ensures every sheet + header row exists.
 * - Creates the Drive folder structure (root > Photos, root > Receipts).
 * - Creates a default admin user (username: admin / password: 12345) if Staff sheet is empty.
 * Safe to run multiple times.
 */
function setupProject() {
  var props = PropertiesService.getScriptProperties();
  var ssId = props.getProperty(CONFIG.PROPS.SPREADSHEET_ID);
  var ss;
  if (!ssId) {
    ss = SpreadsheetApp.create('VALIN CLINIC - Database');
    props.setProperty(CONFIG.PROPS.SPREADSHEET_ID, ss.getId());
  } else {
    ss = SpreadsheetApp.openById(ssId);
  }

  // Remove the default blank "Sheet1" once we have at least one real sheet, keep otherwise.
  Object.keys(CONFIG.SHEETS).forEach(function (key) {
    var name = CONFIG.SHEETS[key];
    ensureSheet_(ss, name, CONFIG.HEADERS[name]);
  });
  var sheet1 = ss.getSheetByName('Sheet1');
  if (sheet1 && ss.getSheets().length > 1) {
    ss.deleteSheet(sheet1);
  }

  // Drive folders
  ensureDriveFolders_();

  // Default admin account
  var staffSheet = getSheet_(CONFIG.SHEETS.STAFF);
  var staffRows = getRows_(staffSheet);
  if (staffRows.length === 0) {
    var salt = Utilities.getUuid();
    appendRow_(staffSheet, CONFIG.HEADERS.Staff, {
      StaffID: 'S' + nextSeq_('STAFF', 4),
      Name: 'ผู้ดูแลระบบ',
      Role: 'ผู้ดูแลระบบ',
      Username: 'admin',
      PasswordHash: hashPassword_('12345', salt),
      PasswordSalt: salt,
      LineID: '',
      Level: 'admin',
      Active: 'TRUE',
      CreatedAt: nowStr_()
    });
  }

  invalidateMasterCache_();

  return { ok: true, spreadsheetId: ss.getId(), spreadsheetUrl: ss.getUrl() };
}

function ensureDriveFolders_() {
  var props = PropertiesService.getScriptProperties();
  var rootId = props.getProperty(CONFIG.PROPS.ROOT_FOLDER_ID);
  var root;
  if (rootId) {
    try { root = DriveApp.getFolderById(rootId); } catch (e) { root = null; }
  }
  if (!root) {
    root = DriveApp.createFolder('VALIN CLINIC - Files');
    props.setProperty(CONFIG.PROPS.ROOT_FOLDER_ID, root.getId());
  }
  var photos = getOrCreateChildFolder_(root, 'Photos');
  var receipts = getOrCreateChildFolder_(root, 'Receipts');
  // Round 17 item 2 — dedicated root for the OPD photo Year/Month/Day tree.
  var opdPhotos = getOrCreateChildFolder_(root, 'OPD Photos');
  props.setProperty(CONFIG.PROPS.PHOTOS_FOLDER_ID, photos.getId());
  props.setProperty(CONFIG.PROPS.RECEIPTS_FOLDER_ID, receipts.getId());
  props.setProperty(CONFIG.PROPS.OPD_PHOTOS_FOLDER_ID, opdPhotos.getId());
  return { rootId: root.getId(), photosId: photos.getId(), receiptsId: receipts.getId(), opdPhotosId: opdPhotos.getId() };
}

function getOrCreateChildFolder_(parent, name) {
  var it = parent.getFoldersByName(name);
  if (it.hasNext()) return it.next();
  return parent.createFolder(name);
}

/** Year > Month(01-12, Thai label optional) > Day folder chain under Receipts root. */
function getReceiptDayFolder_(dateObj) {
  var props = PropertiesService.getScriptProperties();
  var rootId = props.getProperty(CONFIG.PROPS.RECEIPTS_FOLDER_ID);
  if (!rootId) rootId = ensureDriveFolders_().receiptsId;
  var root = DriveApp.getFolderById(rootId);
  var y = String(dateObj.getFullYear());
  var m = ('0' + (dateObj.getMonth() + 1)).slice(-2);
  var d = ('0' + dateObj.getDate()).slice(-2);
  var yFolder = getOrCreateChildFolder_(root, y);
  var mFolder = getOrCreateChildFolder_(yFolder, m);
  var dFolder = getOrCreateChildFolder_(mFolder, d);
  return dFolder;
}

function getPhotosFolder_() {
  var props = PropertiesService.getScriptProperties();
  var id = props.getProperty(CONFIG.PROPS.PHOTOS_FOLDER_ID);
  if (!id) id = ensureDriveFolders_().photosId;
  return DriveApp.getFolderById(id);
}

/** Year > Month(01-12) > Day folder chain under "OPD Photos" root — same
 *  pattern as getReceiptDayFolder_ above (Round 17 item 2). */
function getOpdPhotoDayFolder_(dateObj) {
  var props = PropertiesService.getScriptProperties();
  var rootId = props.getProperty(CONFIG.PROPS.OPD_PHOTOS_FOLDER_ID);
  if (!rootId) rootId = ensureDriveFolders_().opdPhotosId;
  var root = DriveApp.getFolderById(rootId);
  var y = String(dateObj.getFullYear());
  var m = ('0' + (dateObj.getMonth() + 1)).slice(-2);
  var d = ('0' + dateObj.getDate()).slice(-2);
  var yFolder = getOrCreateChildFolder_(root, y);
  var mFolder = getOrCreateChildFolder_(yFolder, m);
  var dFolder = getOrCreateChildFolder_(mFolder, d);
  return dFolder;
}

/* ======================================================================
   UTIL — dates, phone numbers, ids
   ====================================================================== */

function todayStr_() {
  return Utilities.formatDate(new Date(), Session.getScriptTimeZone() || 'Asia/Bangkok', 'yyyy-MM-dd');
}
function nowStr_() {
  return Utilities.formatDate(new Date(), Session.getScriptTimeZone() || 'Asia/Bangkok', 'yyyy-MM-dd HH:mm:ss');
}
function dateStr_(d) {
  if (!d) return '';
  if (typeof d === 'string') return d.length >= 10 ? d.substring(0, 10) : d;
  return Utilities.formatDate(d, Session.getScriptTimeZone() || 'Asia/Bangkok', 'yyyy-MM-dd');
}

/** Keep phone numbers as plain text with leading zeros intact. Strips spaces/dashes. */
function normalizePhone_(phone) {
  if (phone === null || phone === undefined) return '';
  var s = String(phone).trim().replace(/[^\d]/g, '');
  return s;
}

function newId_(prefix) {
  return prefix + Utilities.formatDate(new Date(), Session.getScriptTimeZone() || 'Asia/Bangkok', 'yyMMdd') +
    '-' + Math.floor(Math.random() * 9000 + 1000);
}

/**
 * Sequential, collision-safe receipt number generator: REC-YYYYMMDD-XXX (XXX resets daily).
 * Protected by LockService so concurrent submissions never collide.
 */
function generateReceiptNo_() {
  var lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    var sheet = getSheet_(CONFIG.SHEETS.COUNTERS);
    var today = Utilities.formatDate(new Date(), Session.getScriptTimeZone() || 'Asia/Bangkok', 'yyyyMMdd');
    var rows = getRows_(sheet);
    var found = null;
    for (var i = 0; i < rows.length; i++) {
      if (rows[i].Key === 'RECEIPT_' + today) { found = rows[i]; break; }
    }
    var next;
    if (found) {
      next = Number(found.LastSeq) + 1;
      updateRowByIndex_(sheet, CONFIG.HEADERS.Counters, found._row, { Key: found.Key, LastSeq: next });
    } else {
      next = 1;
      appendRow_(sheet, CONFIG.HEADERS.Counters, { Key: 'RECEIPT_' + today, LastSeq: next });
    }
    var seq = ('00' + next).slice(-3);
    return 'REC-' + today + '-' + seq;
  } finally {
    lock.releaseLock();
  }
}

/** Generic per-key sequential counter (used for PatientID / RID / StockID / etc if needed). */
function nextSeq_(counterKey, pad) {
  var lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    var sheet = getSheet_(CONFIG.SHEETS.COUNTERS);
    var rows = getRows_(sheet);
    var found = null;
    for (var i = 0; i < rows.length; i++) {
      if (rows[i].Key === counterKey) { found = rows[i]; break; }
    }
    var next;
    if (found) {
      next = Number(found.LastSeq) + 1;
      updateRowByIndex_(sheet, CONFIG.HEADERS.Counters, found._row, { Key: counterKey, LastSeq: next });
    } else {
      next = 1;
      appendRow_(sheet, CONFIG.HEADERS.Counters, { Key: counterKey, LastSeq: next });
    }
    return String(next).padStart(pad || 4, '0');
  } finally {
    lock.releaseLock();
  }
}

/* ======================================================================
   AUTH — 2-level (admin / staff) session-token based
   ====================================================================== */

function hashPassword_(password, salt) {
  var digest = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, String(password) + '::' + String(salt));
  return digest.map(function (b) { return ('0' + (b & 0xFF).toString(16)).slice(-2); }).join('');
}

/**
 * login(username, password) -> { ok, token, user:{staffId,name,level,username} } or { ok:false, message }
 * Called from client on the login screen.
 */
function login(username, password) {
  try {
    username = String(username || '').trim();
    var sheet = getSheet_(CONFIG.SHEETS.STAFF);
    var rows = getRows_(sheet);
    var user = null;
    for (var i = 0; i < rows.length; i++) {
      if (String(rows[i].Username).toLowerCase() === username.toLowerCase()) { user = rows[i]; break; }
    }
    if (!user) return { ok: false, message: 'ไม่พบชื่อผู้ใช้นี้ในระบบ' };
    if (String(user.Active).toUpperCase() === 'FALSE') return { ok: false, message: 'บัญชีนี้ถูกระงับการใช้งาน' };
    var hash = hashPassword_(password, user.PasswordSalt);
    if (hash !== user.PasswordHash) return { ok: false, message: 'รหัสผ่านไม่ถูกต้อง' };

    var token = Utilities.getUuid();
    var cache = CacheService.getScriptCache();
    var sessionObj = { staffId: user.StaffID, name: user.Name, level: user.Level, username: user.Username };
    cache.put(CONFIG.CACHE.SESSION_PREFIX + token, JSON.stringify(sessionObj), CONFIG.CACHE.SESSION_TTL);

    return { ok: true, token: token, user: sessionObj };
  } catch (err) {
    return { ok: false, message: 'เกิดข้อผิดพลาด: ' + err.message };
  }
}

function logout(token) {
  try {
    CacheService.getScriptCache().remove(CONFIG.CACHE.SESSION_PREFIX + token);
    return { ok: true };
  } catch (err) {
    return { ok: false, message: err.message };
  }
}

/**
 * Validates a session token and (optionally) enforces a minimum access level.
 * level: 'staff' (any logged-in user) or 'admin' (admin only). Throws on failure —
 * callers should let this bubble up to withFailureHandler on the client.
 */
function requireAuth_(token, level) {
  if (!token) throw new Error('AUTH: กรุณาเข้าสู่ระบบ');
  var cache = CacheService.getScriptCache();
  var raw = cache.get(CONFIG.CACHE.SESSION_PREFIX + token);
  if (!raw) throw new Error('AUTH: เซสชันหมดอายุ กรุณาเข้าสู่ระบบใหม่');
  var session = JSON.parse(raw);
  if (level === 'admin' && session.level !== 'admin') {
    throw new Error('AUTH: การดำเนินการนี้ต้องใช้สิทธิ์ผู้ดูแลระบบ');
  }
  // sliding expiration
  cache.put(CONFIG.CACHE.SESSION_PREFIX + token, raw, CONFIG.CACHE.SESSION_TTL);
  return session;
}

function getCurrentUser(token) {
  var session = requireAuth_(token, 'staff');
  return { ok: true, user: session };
}

function changePassword(token, currentPassword, newPassword) {
  var session = requireAuth_(token, 'staff');
  var sheet = getSheet_(CONFIG.SHEETS.STAFF);
  var rowIndex = findRowIndexById_(sheet, 'StaffID', session.staffId);
  if (rowIndex < 0) return { ok: false, message: 'ไม่พบผู้ใช้' };
  var rows = getRows_(sheet);
  var user = rows.filter(function (r) { return r._row === rowIndex; })[0];
  var currentHash = hashPassword_(currentPassword, user.PasswordSalt);
  if (currentHash !== user.PasswordHash) return { ok: false, message: 'รหัสผ่านเดิมไม่ถูกต้อง' };
  var salt = Utilities.getUuid();
  updateRowByIndex_(sheet, CONFIG.HEADERS.Staff, rowIndex, Object.assign({}, user, {
    PasswordHash: hashPassword_(newPassword, salt), PasswordSalt: salt
  }));
  return { ok: true };
}

/** Admin-only: reset a staff member's password back to the default 12345. */
function resetStaffPassword(token, staffId) {
  requireAuth_(token, 'admin');
  var sheet = getSheet_(CONFIG.SHEETS.STAFF);
  var rowIndex = findRowIndexById_(sheet, 'StaffID', staffId);
  if (rowIndex < 0) return { ok: false, message: 'ไม่พบผู้ใช้' };
  var rows = getRows_(sheet);
  var user = rows.filter(function (r) { return r._row === rowIndex; })[0];
  var salt = Utilities.getUuid();
  updateRowByIndex_(sheet, CONFIG.HEADERS.Staff, rowIndex, Object.assign({}, user, {
    PasswordHash: hashPassword_('12345', salt), PasswordSalt: salt
  }));
  invalidateMasterCache_();
  return { ok: true, message: 'รีเซ็ตรหัสผ่านเป็น 12345 เรียบร้อยแล้ว' };
}

/* ======================================================================
   MASTER CACHE — dropdown / reference data, refreshed on every relevant edit
   ====================================================================== */

/**
 * Returns cached master/reference data used to populate dropdowns across the app:
 * doctors, sales staff, treatment categories, stock categories, clinic info,
 * and (admin only) the staff list. Cached in CacheService for CONFIG.CACHE.MASTER_DATA_TTL
 * seconds; invalidated by invalidateMasterCache_() whenever underlying sheets change.
 */
function getMasterData(token) {
  var session = requireAuth_(token, 'staff');
  var cache = CacheService.getScriptCache();
  var raw = cache.get(CONFIG.CACHE.MASTER_DATA_KEY);
  var data;
  if (raw) {
    data = JSON.parse(raw);
  } else {
    data = buildMasterData_();
    cache.put(CONFIG.CACHE.MASTER_DATA_KEY, JSON.stringify(data), CONFIG.CACHE.MASTER_DATA_TTL);
  }
  var result = Object.assign({}, data);
  if (session.level !== 'admin') delete result.staff; // staff list is admin-only
  return { ok: true, data: result };
}

// Doctors/Sales use a 'Status' column with values 'active' | 'inactive'.
function isActiveStatus_(v) {
  var s = String(v || '').trim().toLowerCase();
  return s === '' || s === 'active' || s === 'ปกติ';
}

function buildMasterData_() {
  var doctors = getRows_(getSheet_(CONFIG.SHEETS.DOCTORS)).filter(function (d) { return isActiveStatus_(d.Status); });
  var salesRows = getRows_(getSheet_(CONFIG.SHEETS.SALES)).filter(function (s) { return isActiveStatus_(s.Status); });
  var staff = getRows_(getSheet_(CONFIG.SHEETS.STAFF));
  var categories = getRows_(getSheet_(CONFIG.SHEETS.CATEGORIES));
  // Round 13: active promotions are now needed by the record modal's
  // "รายการโปร" picker, not just the standalone Promotions page — expose them
  // here so they load with the rest of master data on first app load instead
  // of requiring a separate listPromotions() round trip before the modal
  // can render its dropdown.
  var promoRows = getRows_(getSheet_(CONFIG.SHEETS.PROMOTIONS)).filter(function (p) { return isActiveStatus_(p.Status); });
  var treatmentCats = categories.filter(function (c) { return c.Type === 'treatment'; });
  var stockCats = categories.filter(function (c) { return c.Type === 'stock'; });
  var expenseCats = categories.filter(function (c) { return c.Type === 'expense'; });
  // ClinicSettings is a generic Key/Value store (one row per setting) -> flatten to a single object.
  var settingsRows = getRows_(getSheet_(CONFIG.SHEETS.CLINIC_SETTINGS));
  var clinicInfo = {};
  settingsRows.forEach(function (r) { clinicInfo[r.Key] = r.Value; });

  return {
    doctors: doctors.map(function (d) { return { id: d.DoctorID, name: d.Name, minAmount: d.MinAmount || 0, commissionPct: d.CommissionPct || 0 }; }),
    sales: salesRows.map(function (s) { return { id: s.SalesID, name: s.Name, commissionRate: s.CommissionRate || 0, minAmount: s.MinAmount || 0 }; }),
    staff: staff.map(function (s) { return { id: s.StaffID, name: s.Name, username: s.Username, level: s.Level, active: s.Active }; }),
    treatmentCategories: treatmentCats.map(function (c) { return { id: c.CategoryID, name: c.Name }; }),
    stockCategories: stockCats.map(function (c) { return { id: c.CategoryID, name: c.Name }; }),
    expenseCategories: expenseCats.map(function (c) { return { id: c.CategoryID, name: c.Name }; }),
    promotions: promoRows.map(function (p) { return { id: p.PromoID, name: p.Name, price: Number(p.Price) || 0, doctorSharePct: Number(p.DoctorSharePct) || 0 }; }),
    clinicInfo: clinicInfo
  };
}

function invalidateMasterCache_() {
  CacheService.getScriptCache().remove(CONFIG.CACHE.MASTER_DATA_KEY);
}

/**
 * Full (lightweight) patient / stock lists used to seed client-side
 * suggestion caches on first app load. The client keeps these in memory and
 * filters locally instead of round-tripping to the server on every
 * keystroke; it only re-fetches when told to (after a save/delete of the
 * relevant entity — see invalidatePatientsLiteCache_/invalidateStockLiteCache_
 * callers). Cached server-side too so repeated logins/tabs stay cheap.
 * NOTE: CacheService values are capped at ~100KB — fine for typical clinic
 * volumes, but a very large patient/stock list could exceed it; in that case
 * this simply falls through to reading the sheet fresh every call (the
 * cache.put is wrapped defensively).
 */
function listAllPatientsLite(token) {
  return safeCall_(function () {
    requireAuth_(token, 'staff');
    var cache = CacheService.getScriptCache();
    var raw = cache.get(CONFIG.CACHE.PATIENTS_LITE_KEY);
    if (raw) return { ok: true, data: JSON.parse(raw) };
    var rows = getRows_(getSheet_(CONFIG.SHEETS.PATIENTS));
    var data = rows.map(function (p) { return { id: p.ID, firstName: p.FirstName, lastName: p.LastName, nickname: p.Nickname || '', phone: p.Phone }; });
    try { cache.put(CONFIG.CACHE.PATIENTS_LITE_KEY, JSON.stringify(data), CONFIG.CACHE.LITE_TTL); } catch (e) { /* too large for cache — still return live data */ }
    return { ok: true, data: data };
  });
}
function invalidatePatientsLiteCache_() {
  CacheService.getScriptCache().remove(CONFIG.CACHE.PATIENTS_LITE_KEY);
}

function listAllStockLite(token) {
  return safeCall_(function () {
    requireAuth_(token, 'staff');
    var cache = CacheService.getScriptCache();
    var raw = cache.get(CONFIG.CACHE.STOCK_LITE_KEY);
    if (raw) return { ok: true, data: JSON.parse(raw) };
    var rows = getRows_(getSheet_(CONFIG.SHEETS.STOCK));
    var data = rows.map(function (s) {
      return { id: s.ItemID, name: s.Name, category: s.Category, unit: s.Unit, type: stockTypeOf_(s), prices: STOCK_PRICE_TIERS.map(function (t) { return numOrNull_(s['Price' + t]); }) };
    });
    try { cache.put(CONFIG.CACHE.STOCK_LITE_KEY, JSON.stringify(data), CONFIG.CACHE.LITE_TTL); } catch (e) { /* too large for cache — still return live data */ }
    return { ok: true, data: data };
  });
}
function invalidateStockLiteCache_() {
  CacheService.getScriptCache().remove(CONFIG.CACHE.STOCK_LITE_KEY);
}

/* ======================================================================
   VALIDATION HELPERS — server-side "layer validation"
   ====================================================================== */

function ValidationError_(message) {
  var e = new Error('VALIDATE: ' + message);
  e.isValidation = true;
  return e;
}

/** Throws ValidationError_ listing every missing/blank required field (Thai labels). */
function requireFields_(obj, fieldLabelMap) {
  var missing = [];
  Object.keys(fieldLabelMap).forEach(function (key) {
    var v = obj[key];
    if (v === undefined || v === null || String(v).trim() === '') missing.push(fieldLabelMap[key]);
  });
  if (missing.length) throw ValidationError_('กรุณากรอกข้อมูลให้ครบถ้วน: ' + missing.join(', '));
}

function requireDateStr_(value, label) {
  if (!/^\d{4}-\d{2}-\d{2}$/.test(String(value || ''))) {
    throw ValidationError_(label + ' ต้องอยู่ในรูปแบบ yyyy-mm-dd');
  }
}

function requireNumber_(value, label, opts) {
  opts = opts || {};
  var n = Number(value);
  if (isNaN(n)) throw ValidationError_(label + ' ต้องเป็นตัวเลข');
  if (opts.min !== undefined && n < opts.min) throw ValidationError_(label + ' ต้องไม่น้อยกว่า ' + opts.min);
  return n;
}

/** Wraps a server function so validation/auth errors return {ok:false,message} instead of throwing raw to the client. */
function safeCall_(fn) {
  try {
    return fn();
  } catch (err) {
    var msg = String(err.message || err);
    if (msg.indexOf('VALIDATE: ') === 0) msg = msg.substring('VALIDATE: '.length);
    else if (msg.indexOf('AUTH: ') === 0) msg = msg.substring('AUTH: '.length);
    return { ok: false, message: msg };
  }
}

/** Only keeps keys with non-empty values — used so filter objects only apply the filters actually set. */
function compactFilters_(filters) {
  var out = {};
  if (!filters) return out;
  Object.keys(filters).forEach(function (k) {
    var v = filters[k];
    if (v !== undefined && v !== null && String(v).trim() !== '') out[k] = v;
  });
  return out;
}

function writeAudit_(session, action, entity, entityId, detail) {
  try {
    var sheet = getSheet_(CONFIG.SHEETS.AUDIT_LOG);
    appendRow_(sheet, CONFIG.HEADERS.AuditLog, {
      Timestamp: nowStr_(), User: session ? session.username : '', Action: action,
      Entity: entity, EntityID: entityId, Detail: detail || ''
    });
  } catch (e) { /* audit failures must never block the primary operation */ }
}

/* ======================================================================
   PATIENTS — ข้อมูลคนไข้
   ====================================================================== */

function listPatients(token, opts) {
  return safeCall_(function () {
    requireAuth_(token, 'staff');
    opts = opts || {};
    var filters = compactFilters_(opts.filters);
    var sheet = getSheet_(CONFIG.SHEETS.PATIENTS);
    var rows = getRows_(sheet);
    if (filters.q) {
      var q = String(filters.q).trim().toLowerCase();
      var qDigits = q.replace(/[^\d]/g, '');
      rows = rows.filter(function (p) {
        var fullName = (p.FirstName + ' ' + p.LastName).toLowerCase();
        return fullName.indexOf(q) > -1 || (qDigits && String(p.Phone).indexOf(qDigits) > -1);
      });
    }
    rows.sort(function (a, b) { return String(b.CreatedAt).localeCompare(String(a.CreatedAt)); });
    var page = paginate_(rows, opts.page, opts.pageSize);
    page.items = page.items.map(mapPatientOut_);
    return { ok: true, data: page };
  });
}

function mapPatientOut_(p) {
  return {
    id: p.ID, firstName: p.FirstName, lastName: p.LastName, nickname: p.Nickname || '', phone: p.Phone, dob: dateStr_(p.DOB),
    gender: p.Gender, email: p.Email, photoFileId: p.PhotoFileId, joinedDate: dateStr_(p.JoinedDate), row: p._row
  };
}

/**
 * Read-only preview of the next auto-generated patient ID, shown in the
 * "new patient" modal before the record is actually saved. This does NOT
 * reserve/increment the counter (savePatient's nextSeq_('PATIENT',5) call
 * still does that at save time) — it's a cosmetic preview only, so on the
 * rare case of two staff opening the modal at the same moment the shown ID
 * could shift by the time either one actually saves. That's an acceptable
 * trade-off since the ID is never user-editable either way.
 */
function getNextPatientId(token) {
  return safeCall_(function () {
    requireAuth_(token, 'staff');
    return { ok: true, data: 'P' + peekNextSeq_('PATIENT', 5) };
  });
}
function peekNextSeq_(counterKey, pad) {
  var rows = getRows_(getSheet_(CONFIG.SHEETS.COUNTERS));
  var found = rows.filter(function (r) { return r.Key === counterKey; })[0];
  var next = found ? Number(found.LastSeq) + 1 : 1;
  return String(next).padStart(pad || 4, '0');
}

/**
 * Patient profile bundle for the right-side slide-out panel: the patient
 * record plus computed stats (visit count, cumulative spend excluding
 * cancelled visits) and recent treatment history.
 */
function getPatientProfile(token, id) {
  return safeCall_(function () {
    requireAuth_(token, 'staff');
    var rows = getRows_(getSheet_(CONFIG.SHEETS.PATIENTS));
    var p = rows.filter(function (r) { return r.ID === id; })[0];
    if (!p) throw ValidationError_('ไม่พบข้อมูลคนไข้');
    var records = getRows_(getSheet_(CONFIG.SHEETS.RECORDS)).filter(function (r) { return r.PatientID === id; });
    // Same date-sort fix as listRecords (round 12) — compare normalized date
    // strings, not raw Date objects.
    records.sort(function (a, b) { return dateStr_(b.Date).localeCompare(dateStr_(a.Date)) || String(b.CreatedAt).localeCompare(String(a.CreatedAt)); });
    var visitCount = records.length;
    var totalSpent = records.reduce(function (s, r) {
      if (r.Status === 'cancelled') return s;
      var useCanula = String(r.UseCanula) === 'TRUE' || r.UseCanula === true;
      return s + (Number(r.Amount) || 0) + (Number(r.PromoAmount) || 0) + (useCanula ? (Number(r.CanulaAmount) || 0) : 0) + (Number(r.MedicineAmount) || 0) + (Number(r.OtherAmount) || 0);
    }, 0);
    return {
      ok: true,
      data: {
        patient: mapPatientOut_(p),
        visitCount: visitCount,
        totalSpent: totalSpent,
        records: records.slice(0, 40).map(mapRecordOut_)
      }
    };
  });
}

/** Lightweight suggest list for search-as-you-type widgets (used by Records/Receipt patient search too). */
function searchPatients(token, q, limit) {
  return safeCall_(function () {
    requireAuth_(token, 'staff');
    q = String(q || '').trim().toLowerCase();
    if (!q) return { ok: true, data: [] };
    var qDigits = q.replace(/[^\d]/g, '');
    var rows = getRows_(getSheet_(CONFIG.SHEETS.PATIENTS));
    var matches = rows.filter(function (p) {
      var fullName = (p.FirstName + ' ' + p.LastName).toLowerCase();
      return fullName.indexOf(q) > -1 || (qDigits && String(p.Phone).indexOf(qDigits) > -1);
    }).slice(0, limit || 8);
    return { ok: true, data: matches.map(mapPatientOut_) };
  });
}

function getPatient(token, id) {
  return safeCall_(function () {
    requireAuth_(token, 'staff');
    var rows = getRows_(getSheet_(CONFIG.SHEETS.PATIENTS));
    var p = rows.filter(function (r) { return r.ID === id; })[0];
    if (!p) throw ValidationError_('ไม่พบข้อมูลคนไข้');
    return { ok: true, data: mapPatientOut_(p) };
  });
}

function savePatient(token, patient) {
  return safeCall_(function () {
    var session = requireAuth_(token, 'staff');
    requireFields_(patient, { firstName: 'ชื่อ', lastName: 'นามสกุล', phone: 'เบอร์โทร' });
    var phone = normalizePhone_(patient.phone);
    if (phone.length < 9 || phone.length > 10) throw ValidationError_('เบอร์โทรไม่ถูกต้อง (9-10 หลัก)');
    if (patient.dob) requireDateStr_(patient.dob, 'วันเกิด');

    var lock = LockService.getScriptLock();
    lock.waitLock(30000);
    try {
      var sheet = getSheet_(CONFIG.SHEETS.PATIENTS);
      var isNew = !patient.id;
      var obj = {
        ID: isNew ? ('P' + nextSeq_('PATIENT', 5)) : patient.id,
        FirstName: patient.firstName, LastName: patient.lastName, Nickname: patient.nickname || '', Phone: phone,
        DOB: patient.dob || '', Gender: patient.gender || '', Email: patient.email || '',
        PhotoFileId: patient.photoFileId || '', JoinedDate: patient.joinedDate || todayStr_(),
        CreatedAt: isNew ? nowStr_() : patient.createdAt || nowStr_(), UpdatedAt: nowStr_()
      };
      if (isNew) {
        appendRow_(sheet, CONFIG.HEADERS.Patients, obj);
      } else {
        var rowIndex = findRowIndexById_(sheet, 'ID', patient.id);
        if (rowIndex < 0) throw ValidationError_('ไม่พบข้อมูลคนไข้ที่จะแก้ไข');
        updateRowByIndex_(sheet, CONFIG.HEADERS.Patients, rowIndex, obj);
      }
      writeAudit_(session, isNew ? 'CREATE' : 'UPDATE', 'Patient', obj.ID, obj.FirstName + ' ' + obj.LastName);
      invalidatePatientsLiteCache_();
      return { ok: true, data: mapPatientOut_(Object.assign({ _row: 0 }, obj)) };
    } finally {
      lock.releaseLock();
    }
  });
}

function deletePatient(token, id) {
  return safeCall_(function () {
    var session = requireAuth_(token, 'admin');
    var sheet = getSheet_(CONFIG.SHEETS.PATIENTS);
    var rowIndex = findRowIndexById_(sheet, 'ID', id);
    if (rowIndex < 0) throw ValidationError_('ไม่พบข้อมูลคนไข้');
    deleteRowByIndex_(sheet, rowIndex);
    writeAudit_(session, 'DELETE', 'Patient', id, '');
    invalidatePatientsLiteCache_();
    return { ok: true };
  });
}

/* ======================================================================
   ตั้งค่าคลีนิค — Doctors / Sales / Categories / Staff / ClinicSettings
   All writes here are admin-only and invalidate the master-data cache.
   ====================================================================== */

function listDoctors(token) {
  return safeCall_(function () {
    requireAuth_(token, 'staff');
    var rows = getRows_(getSheet_(CONFIG.SHEETS.DOCTORS));
    return { ok: true, data: rows.map(function (d) {
      return { id: d.DoctorID, name: d.Name, status: d.Status || 'active', minAmount: d.MinAmount || 0, commissionPct: d.CommissionPct || 0, row: d._row };
    }) };
  });
}

function saveDoctor(token, doctor) {
  return safeCall_(function () {
    var session = requireAuth_(token, 'admin');
    requireFields_(doctor, { name: 'ชื่อแพทย์' });
    var sheet = getSheet_(CONFIG.SHEETS.DOCTORS);
    var isNew = !doctor.id;
    var obj = {
      DoctorID: isNew ? ('D' + nextSeq_('DOCTOR', 3)) : doctor.id,
      Name: doctor.name, Status: doctor.status || 'active', MinAmount: requireNumber_(doctor.minAmount || 0, 'ยอดขั้นต่ำ', { min: 0 }),
      CommissionPct: requireNumber_(doctor.commissionPct || 0, '% Commission', { min: 0 }),
      CreatedAt: isNew ? nowStr_() : (doctor.createdAt || nowStr_())
    };
    if (isNew) appendRow_(sheet, CONFIG.HEADERS.Doctors, obj);
    else {
      var rowIndex = findRowIndexById_(sheet, 'DoctorID', doctor.id);
      if (rowIndex < 0) throw ValidationError_('ไม่พบข้อมูลแพทย์');
      updateRowByIndex_(sheet, CONFIG.HEADERS.Doctors, rowIndex, obj);
    }
    invalidateMasterCache_();
    writeAudit_(session, isNew ? 'CREATE' : 'UPDATE', 'Doctor', obj.DoctorID, obj.Name);
    return { ok: true, data: obj };
  });
}

function deleteDoctor(token, id) {
  return safeCall_(function () {
    var session = requireAuth_(token, 'admin');
    var sheet = getSheet_(CONFIG.SHEETS.DOCTORS);
    var rowIndex = findRowIndexById_(sheet, 'DoctorID', id);
    if (rowIndex < 0) throw ValidationError_('ไม่พบข้อมูลแพทย์');
    deleteRowByIndex_(sheet, rowIndex);
    invalidateMasterCache_();
    writeAudit_(session, 'DELETE', 'Doctor', id, '');
    return { ok: true };
  });
}

function listSales(token) {
  return safeCall_(function () {
    requireAuth_(token, 'staff');
    var rows = getRows_(getSheet_(CONFIG.SHEETS.SALES));
    return { ok: true, data: rows.map(function (s) {
      return { id: s.SalesID, name: s.Name, status: s.Status || 'active', commissionRate: s.CommissionRate || 0, minAmount: s.MinAmount || 0, row: s._row };
    }) };
  });
}

function saveSales(token, sales) {
  return safeCall_(function () {
    var session = requireAuth_(token, 'admin');
    requireFields_(sales, { name: 'ชื่อเซลล์' });
    var sheet = getSheet_(CONFIG.SHEETS.SALES);
    var isNew = !sales.id;
    var obj = {
      SalesID: isNew ? ('SL' + nextSeq_('SALES', 3)) : sales.id,
      Name: sales.name, Status: sales.status || 'active',
      CommissionRate: requireNumber_(sales.commissionRate || 0, 'อัตราค่าคอมมิชชั่น', { min: 0 }),
      MinAmount: requireNumber_(sales.minAmount || 0, 'ยอดขั้นต่ำ', { min: 0 }),
      CreatedAt: isNew ? nowStr_() : (sales.createdAt || nowStr_())
    };
    if (isNew) appendRow_(sheet, CONFIG.HEADERS.Sales, obj);
    else {
      var rowIndex = findRowIndexById_(sheet, 'SalesID', sales.id);
      if (rowIndex < 0) throw ValidationError_('ไม่พบข้อมูลเซลล์');
      updateRowByIndex_(sheet, CONFIG.HEADERS.Sales, rowIndex, obj);
    }
    invalidateMasterCache_();
    writeAudit_(session, isNew ? 'CREATE' : 'UPDATE', 'Sales', obj.SalesID, obj.Name);
    return { ok: true, data: obj };
  });
}

function deleteSales(token, id) {
  return safeCall_(function () {
    var session = requireAuth_(token, 'admin');
    var sheet = getSheet_(CONFIG.SHEETS.SALES);
    var rowIndex = findRowIndexById_(sheet, 'SalesID', id);
    if (rowIndex < 0) throw ValidationError_('ไม่พบข้อมูลเซลล์');
    deleteRowByIndex_(sheet, rowIndex);
    invalidateMasterCache_();
    writeAudit_(session, 'DELETE', 'Sales', id, '');
    return { ok: true };
  });
}

/* ---------------------------------------------------------------- */
/* PROMOTIONS — จัดการโปรโมชันหน้าคลินิก. Since round 13 this list is also    */
/* consumed live by the treatment record modal's "รายการโปร" picker (via     */
/* buildMasterData_) and its DoctorSharePct feeds the revenue-split engine — */
/* see the Promotions header comment above. Every write here must           */
/* invalidateMasterCache_() so record-modal dropdowns/pricing don't go       */
/* stale against the (up to MASTER_DATA_TTL-seconds-old) cached master data. */
/* ---------------------------------------------------------------- */
function listPromotions(token) {
  return safeCall_(function () {
    requireAuth_(token, 'staff');
    var rows = getRows_(getSheet_(CONFIG.SHEETS.PROMOTIONS));
    return { ok: true, data: rows.map(function (p) {
      var doctorPct = Number(p.DoctorSharePct) || 0;
      return {
        id: p.PromoID, name: p.Name, price: Number(p.Price) || 0,
        doctorSharePct: doctorPct, clinicSharePct: 100 - doctorPct,
        status: p.Status || 'active'
      };
    }) };
  });
}

function savePromotion(token, promo) {
  return safeCall_(function () {
    var session = requireAuth_(token, 'admin');
    requireFields_(promo, { name: 'ชื่อโปรโมชัน' });
    var sheet = getSheet_(CONFIG.SHEETS.PROMOTIONS);
    var isNew = !promo.id;
    var doctorPct = requireNumber_(promo.doctorSharePct || 0, '% ส่วนแบ่งแพทย์', { min: 0 });
    if (doctorPct > 100) throw ValidationError_('% ส่วนแบ่งแพทย์ ต้องไม่เกิน 100');
    var obj = {
      PromoID: isNew ? ('PM' + nextSeq_('PROMO', 3)) : promo.id,
      Name: promo.name,
      Price: requireNumber_(promo.price || 0, 'ราคา', { min: 0 }),
      DoctorSharePct: doctorPct,
      Status: promo.status || 'active',
      CreatedAt: isNew ? nowStr_() : (promo.createdAt || nowStr_()),
      UpdatedAt: nowStr_()
    };
    if (isNew) appendRow_(sheet, CONFIG.HEADERS.Promotions, obj);
    else {
      var rowIndex = findRowIndexById_(sheet, 'PromoID', promo.id);
      if (rowIndex < 0) throw ValidationError_('ไม่พบข้อมูลโปรโมชัน');
      updateRowByIndex_(sheet, CONFIG.HEADERS.Promotions, rowIndex, obj);
    }
    invalidateMasterCache_();
    writeAudit_(session, isNew ? 'CREATE' : 'UPDATE', 'Promotion', obj.PromoID, obj.Name);
    return { ok: true, data: obj };
  });
}

/** Quick active/inactive flip from the table's toggle switch — avoids making
 * the client resend the full promotion payload just to change one field. */
function togglePromotionStatus(token, id) {
  return safeCall_(function () {
    var session = requireAuth_(token, 'admin');
    var sheet = getSheet_(CONFIG.SHEETS.PROMOTIONS);
    var rowIndex = findRowIndexById_(sheet, 'PromoID', id);
    if (rowIndex < 0) throw ValidationError_('ไม่พบข้อมูลโปรโมชัน');
    var row = getRows_(sheet).filter(function (r) { return r.PromoID === id; })[0];
    var newStatus = row.Status === 'active' ? 'inactive' : 'active';
    // updateRowByIndex_ rewrites the WHOLE row (any header missing from the
    // object it's given is written as blank) — merge onto the existing row
    // rather than passing a partial {Status, UpdatedAt} obj, or every other
    // column (name/price/%) would be wiped out.
    var obj = {
      PromoID: row.PromoID, Name: row.Name, Price: row.Price, DoctorSharePct: row.DoctorSharePct,
      Status: newStatus, CreatedAt: row.CreatedAt, UpdatedAt: nowStr_()
    };
    updateRowByIndex_(sheet, CONFIG.HEADERS.Promotions, rowIndex, obj);
    invalidateMasterCache_();
    writeAudit_(session, 'UPDATE', 'Promotion', id, 'status -> ' + newStatus);
    return { ok: true, data: { status: newStatus } };
  });
}

function deletePromotion(token, id) {
  return safeCall_(function () {
    var session = requireAuth_(token, 'admin');
    var sheet = getSheet_(CONFIG.SHEETS.PROMOTIONS);
    var rowIndex = findRowIndexById_(sheet, 'PromoID', id);
    if (rowIndex < 0) throw ValidationError_('ไม่พบข้อมูลโปรโมชัน');
    deleteRowByIndex_(sheet, rowIndex);
    invalidateMasterCache_();
    writeAudit_(session, 'DELETE', 'Promotion', id, '');
    return { ok: true };
  });
}

/** type: 'treatment' | 'stock' | 'expense' */
function listCategories(token, type) {
  return safeCall_(function () {
    requireAuth_(token, 'staff');
    var rows = getRows_(getSheet_(CONFIG.SHEETS.CATEGORIES));
    if (type) rows = rows.filter(function (c) { return c.Type === type; });
    rows.sort(function (a, b) { return (Number(a.SortOrder) || 0) - (Number(b.SortOrder) || 0); });
    return { ok: true, data: rows.map(function (c) {
      return { id: c.CategoryID, type: c.Type, name: c.Name, sortOrder: c.SortOrder || 0, row: c._row };
    }) };
  });
}

function saveCategory(token, cat) {
  return safeCall_(function () {
    var session = requireAuth_(token, 'admin');
    requireFields_(cat, { name: 'ชื่อหมวดหมู่', type: 'ประเภท' });
    if (['treatment', 'stock', 'expense'].indexOf(cat.type) === -1) throw ValidationError_('ประเภทหมวดหมู่ไม่ถูกต้อง');
    var sheet = getSheet_(CONFIG.SHEETS.CATEGORIES);
    var isNew = !cat.id;
    var obj = {
      CategoryID: isNew ? ('C' + nextSeq_('CATEGORY', 3)) : cat.id,
      Type: cat.type, Name: cat.name, SortOrder: cat.sortOrder || 0
    };
    if (isNew) appendRow_(sheet, CONFIG.HEADERS.Categories, obj);
    else {
      var rowIndex = findRowIndexById_(sheet, 'CategoryID', cat.id);
      if (rowIndex < 0) throw ValidationError_('ไม่พบหมวดหมู่');
      updateRowByIndex_(sheet, CONFIG.HEADERS.Categories, rowIndex, obj);
    }
    invalidateMasterCache_();
    writeAudit_(session, isNew ? 'CREATE' : 'UPDATE', 'Category', obj.CategoryID, obj.Name);
    return { ok: true, data: obj };
  });
}

function deleteCategory(token, id) {
  return safeCall_(function () {
    var session = requireAuth_(token, 'admin');
    var sheet = getSheet_(CONFIG.SHEETS.CATEGORIES);
    var rowIndex = findRowIndexById_(sheet, 'CategoryID', id);
    if (rowIndex < 0) throw ValidationError_('ไม่พบหมวดหมู่');
    deleteRowByIndex_(sheet, rowIndex);
    invalidateMasterCache_();
    writeAudit_(session, 'DELETE', 'Category', id, '');
    return { ok: true };
  });
}

/** Admin-only staff/user list & management (create/edit/deactivate; password reset via resetStaffPassword). */
function listStaff(token) {
  return safeCall_(function () {
    requireAuth_(token, 'admin');
    var rows = getRows_(getSheet_(CONFIG.SHEETS.STAFF));
    return { ok: true, data: rows.map(function (s) {
      return { id: s.StaffID, name: s.Name, role: s.Role, username: s.Username, lineId: s.LineID, level: s.Level, active: s.Active !== 'FALSE', row: s._row };
    }) };
  });
}

function saveStaff(token, staff) {
  return safeCall_(function () {
    var session = requireAuth_(token, 'admin');
    requireFields_(staff, { name: 'ชื่อ', username: 'ชื่อผู้ใช้', level: 'สิทธิ์การเข้าถึง' });
    if (['admin', 'staff'].indexOf(staff.level) === -1) throw ValidationError_('สิทธิ์การเข้าถึงไม่ถูกต้อง');
    var sheet = getSheet_(CONFIG.SHEETS.STAFF);
    var rows = getRows_(sheet);
    var usernameTaken = rows.some(function (s) {
      return String(s.Username).toLowerCase() === String(staff.username).toLowerCase() && s.StaffID !== staff.id;
    });
    if (usernameTaken) throw ValidationError_('ชื่อผู้ใช้นี้ถูกใช้แล้ว');

    var isNew = !staff.id;
    if (isNew) {
      var salt = Utilities.getUuid();
      var obj = {
        StaffID: 'S' + nextSeq_('STAFF', 4), Name: staff.name, Role: staff.role || '', Username: staff.username,
        PasswordHash: hashPassword_('12345', salt), PasswordSalt: salt, LineID: staff.lineId || '',
        Level: staff.level, Active: 'TRUE', CreatedAt: nowStr_()
      };
      appendRow_(sheet, CONFIG.HEADERS.Staff, obj);
      invalidateMasterCache_();
      writeAudit_(session, 'CREATE', 'Staff', obj.StaffID, obj.Name);
      return { ok: true, data: obj, message: 'สร้างผู้ใช้ใหม่แล้ว รหัสผ่านเริ่มต้นคือ 12345' };
    } else {
      var rowIndex = findRowIndexById_(sheet, 'StaffID', staff.id);
      if (rowIndex < 0) throw ValidationError_('ไม่พบผู้ใช้');
      var existing = rows.filter(function (r) { return r._row === rowIndex; })[0];
      var updated = Object.assign({}, existing, {
        Name: staff.name, Role: staff.role || '', Username: staff.username,
        LineID: staff.lineId || '', Level: staff.level, Active: staff.active === false ? 'FALSE' : 'TRUE'
      });
      updateRowByIndex_(sheet, CONFIG.HEADERS.Staff, rowIndex, updated);
      invalidateMasterCache_();
      writeAudit_(session, 'UPDATE', 'Staff', staff.id, staff.name);
      return { ok: true, data: updated };
    }
  });
}

function deleteStaff(token, id) {
  return safeCall_(function () {
    var session = requireAuth_(token, 'admin');
    if (id === session.staffId) throw ValidationError_('ไม่สามารถลบบัญชีของตัวเองได้');
    var sheet = getSheet_(CONFIG.SHEETS.STAFF);
    var rowIndex = findRowIndexById_(sheet, 'StaffID', id);
    if (rowIndex < 0) throw ValidationError_('ไม่พบผู้ใช้');
    deleteRowByIndex_(sheet, rowIndex);
    invalidateMasterCache_();
    writeAudit_(session, 'DELETE', 'Staff', id, '');
    return { ok: true };
  });
}

/** ข้อมูลบริษัท (clinic info for receipts) — stored as ClinicSettings key/value rows. */
function getClinicSettings(token) {
  return safeCall_(function () {
    requireAuth_(token, 'staff');
    var rows = getRows_(getSheet_(CONFIG.SHEETS.CLINIC_SETTINGS));
    var out = {};
    rows.forEach(function (r) { out[r.Key] = r.Value; });
    return { ok: true, data: out };
  });
}

/** settings: plain object of key->value, e.g. {name, address, phone, taxId, logoFileId} */
function saveClinicSettings(token, settings) {
  return safeCall_(function () {
    var session = requireAuth_(token, 'admin');
    var sheet = getSheet_(CONFIG.SHEETS.CLINIC_SETTINGS);
    var rows = getRows_(sheet);
    Object.keys(settings || {}).forEach(function (key) {
      var existing = rows.filter(function (r) { return r.Key === key; })[0];
      if (existing) {
        updateRowByIndex_(sheet, CONFIG.HEADERS.ClinicSettings, existing._row, { Key: key, Value: settings[key] });
      } else {
        appendRow_(sheet, CONFIG.HEADERS.ClinicSettings, { Key: key, Value: settings[key] });
      }
    });
    invalidateMasterCache_();
    writeAudit_(session, 'UPDATE', 'ClinicSettings', '', JSON.stringify(settings));
    return { ok: true };
  });
}

/* ======================================================================
   จัดการสต๊อกสินค้า/อุปกรณ์ — Stock
   ====================================================================== */

// ราคาขาย quantity tiers — one Stock column per tier (Price1, Price3, ...).
var STOCK_PRICE_TIERS = [1, 3, 5, 10, 20, 50];
// StockLog.Kind values. '' (rows written before Kind existed) reads as 'update'.
var STOCK_LOG_KINDS = { update: 1, in: 1, out: 1, set: 1, import: 1, opening: 1, record: 1 };

function listStock(token, opts) {
  return safeCall_(function () {
    var session = requireAuth_(token, 'staff');
    opts = opts || {};
    var filters = compactFilters_(opts.filters);
    var rows = getRows_(getSheet_(CONFIG.SHEETS.STOCK));
    if (filters.q) {
      var q = String(filters.q).trim().toLowerCase();
      rows = rows.filter(function (r) { return String(r.Name).toLowerCase().indexOf(q) > -1; });
    }
    if (filters.category) rows = rows.filter(function (r) { return r.Category === filters.category; });
    if (filters.type) rows = rows.filter(function (r) { return stockTypeOf_(r) === filters.type; });

    // Round 25 item 1: "เรียงตาม update ล่าสุด" — sort by each item's latest
    // StockLog entry's Date (items with no log entries yet sink to the
    // bottom, then alphabetical as a tiebreaker/fallback). Join once here
    // rather than per-item, and hand the same latest-entry object to
    // mapStockOut_ so the row can show the latest จำนวนคงเหลือ without a
    // second round-trip per item.
    var latestByItem = latestStockLogByItem_();
    rows.sort(function (a, b) {
      var la = latestByItem[a.ItemID], lb = latestByItem[b.ItemID];
      var da = la ? dateStr_(la.Date) : '';
      var db = lb ? dateStr_(lb.Date) : '';
      if (da !== db) return db.localeCompare(da);
      return String(a.Name).localeCompare(String(b.Name), 'th');
    });
    var showCost = session.level === 'admin';
    var page = paginate_(rows, opts.page, opts.pageSize);
    page.items = page.items.map(function (r) { return mapStockOut_(r, latestByItem[r.ItemID], showCost); });
    return { ok: true, data: page };
  });
}

/** Returns {ItemID: latestStockLogRow} — the most recent (by Date, then CreatedAt as a same-day tiebreaker) StockLog row per item. Shared by listStock's sort/display and by every StockLog write's Stock.Qty re-sync so "latest wins" is defined in exactly one place. */
function latestStockLogByItem_() {
  var logRows = getRows_(getSheet_(CONFIG.SHEETS.STOCK_LOG));
  var latestByItem = {};
  logRows.forEach(function (l) {
    var cur = latestByItem[l.ItemID];
    if (!cur || dateStr_(l.Date) > dateStr_(cur.Date) ||
      (dateStr_(l.Date) === dateStr_(cur.Date) && String(l.CreatedAt) > String(cur.CreatedAt))) {
      latestByItem[l.ItemID] = l;
    }
  });
  return latestByItem;
}

/** Re-syncs Stock.Qty for one item to its TRUE latest StockLog entry (by date, not necessarily whichever entry was just saved/edited — a backdated entry may not be the latest). Called after every StockLog write so Stock.Qty stays a correct denormalized cache by construction. */
function resyncStockQtyFromLog_(itemId) {
  var stockSheet = getSheet_(CONFIG.SHEETS.STOCK);
  var latest = latestStockLogByItem_()[itemId];
  if (!latest) return;
  var stockRowIndex = findRowIndexById_(stockSheet, 'ItemID', itemId);
  if (stockRowIndex < 0) return;
  var stockRow = getRows_(stockSheet).filter(function (r) { return r._row === stockRowIndex; })[0];
  updateRowByIndex_(stockSheet, CONFIG.HEADERS.Stock, stockRowIndex, Object.assign({}, stockRow, { Qty: Number(latest.QtyRemaining) || 0, UpdatedAt: nowStr_() }));
}

/** 'medicine' (ยา) or 'supply' (เวชภัณฑ์) — rows saved before StockType existed are blank and count as 'supply'. */
function stockTypeOf_(s) {
  return s.StockType === 'medicine' ? 'medicine' : 'supply';
}

/** Blank cell -> null (so the client can show "-"), anything else -> Number. */
function numOrNull_(v) {
  return (v === '' || v === null || v === undefined) ? null : (Number(v) || 0);
}

/** Optional numeric input (blank allowed, else must be >= 0) -> value to store ('' for blank). */
function optionalNumber_(v, label) {
  if (v === '' || v === null || v === undefined) return '';
  return requireNumber_(v, label, { min: 0 });
}

/** `showCost` must only be true for admins — ราคาทุน never leaves the server otherwise. */
function mapStockOut_(s, latestLog, showCost) {
  var out = {
    id: s.ItemID, name: s.Name, category: s.Category, unit: s.Unit, qty: Number(s.Qty) || 0, row: s._row,
    type: stockTypeOf_(s),
    minQty: numOrNull_(s.MinQty),
    prices: STOCK_PRICE_TIERS.map(function (t) { return numOrNull_(s['Price' + t]); }),
    latestDate: latestLog ? dateStr_(latestLog.Date) : null,
    latestQtyRemaining: latestLog ? (Number(latestLog.QtyRemaining) || 0) : null
  };
  if (showCost) out.costPrice = numOrNull_(s.CostPrice);
  return out;
}

function searchStock(token, q, limit) {
  return safeCall_(function () {
    var session = requireAuth_(token, 'staff');
    q = String(q || '').trim().toLowerCase();
    if (!q) return { ok: true, data: [] };
    var rows = getRows_(getSheet_(CONFIG.SHEETS.STOCK));
    var matches = rows.filter(function (r) { return String(r.Name).toLowerCase().indexOf(q) > -1; }).slice(0, limit || 8);
    return { ok: true, data: matches.map(function (r) { return mapStockOut_(r, null, session.level === 'admin'); }) };
  });
}

/** Builds one StockLog row object. qtyIn/qtyOut/qtyRemaining are already-validated numbers. */
function stockLogRow_(session, logId, date, item, kind, qtyIn, qtyOut, qtyRemaining, note) {
  return {
    LogID: logId, Date: date, ItemID: item.ItemID, ItemName: item.Name,
    QtyIn: qtyIn, QtyRemaining: qtyRemaining, Note: note || '',
    CreatedBy: session.username, CreatedAt: nowStr_(), QtyOut: qtyOut, Kind: kind
  };
}

/**
 * item: {id?, type, name, category, unit, minQty, costPrice, prices:[6], qty}.
 * `qty` is only used when creating an item (written as a ยอดยกมา StockLog
 * entry); an existing item's quantity only changes through StockLog writes
 * (adjustStock / saveStockUpdate / importStockCsv), never through this form.
 */
function saveStockItem(token, item) {
  return safeCall_(function () {
    var session = requireAuth_(token, 'staff');
    requireFields_(item, { name: 'ชื่อสินค้า', category: 'หมวดหมู่', unit: 'หน่วยนับ' });
    var isAdmin = session.level === 'admin';
    var prices = Array.isArray(item.prices) ? item.prices : [];

    var lock = LockService.getScriptLock();
    lock.waitLock(30000);
    try {
      var sheet = getSheet_(CONFIG.SHEETS.STOCK);
      var isNew = !item.id;
      var existing = null;
      if (!isNew) {
        existing = getRows_(sheet).filter(function (r) { return r.ItemID === item.id; })[0];
        if (!existing) throw ValidationError_('ไม่พบสินค้า');
      }
      var initialQty = isNew ? requireNumber_(item.qty || 0, 'จำนวนคงเหลือ', { min: 0 }) : 0;
      var obj = Object.assign({}, existing || {}, {
        ItemID: isNew ? ('ST' + nextSeq_('STOCK', 4)) : item.id,
        Name: item.name, Category: item.category, Unit: item.unit,
        Qty: isNew ? initialQty : (Number(existing.Qty) || 0),
        CreatedAt: isNew ? nowStr_() : existing.CreatedAt, UpdatedAt: nowStr_(),
        StockType: item.type === 'medicine' ? 'medicine' : 'supply',
        MinQty: optionalNumber_(item.minQty, 'จำนวนสต๊อกขั้นต่ำ'),
        CostPrice: isAdmin ? optionalNumber_(item.costPrice, 'ราคาทุน') : (existing ? existing.CostPrice : '')
      });
      STOCK_PRICE_TIERS.forEach(function (t, i) { obj['Price' + t] = optionalNumber_(prices[i], 'ราคาขาย ' + t + ' ชิ้น'); });
      delete obj._row;

      if (isNew) {
        appendRow_(sheet, CONFIG.HEADERS.Stock, obj);
        if (initialQty > 0) {
          appendRow_(getSheet_(CONFIG.SHEETS.STOCK_LOG), CONFIG.HEADERS.StockLog,
            stockLogRow_(session, 'SL' + nextSeq_('STOCKLOG', 5), todayStr_(), obj, 'opening', initialQty, 0, initialQty, 'ยอดยกมา'));
        }
      } else {
        updateRowByIndex_(sheet, CONFIG.HEADERS.Stock, existing._row, obj);
      }
      writeAudit_(session, isNew ? 'CREATE' : 'UPDATE', 'Stock', obj.ItemID, obj.Name);
      invalidateStockLiteCache_();
      return { ok: true, data: mapStockOut_(Object.assign({ _row: 0 }, obj), null, isAdmin) };
    } finally {
      lock.releaseLock();
    }
  });
}

/**
 * Adjusts stock quantity by a signed delta. Previously called automatically
 * from saveRecord/deleteRecord to deduct/restore supplies used on a
 * treatment record; Round 14 item 1.2 removed that automatic linkage.
 * Left in place, currently unused (manual adjustments go through
 * adjustStock below, which also writes the StockLog history).
 */
function adjustStockQty_(itemId, delta) {
  var sheet = getSheet_(CONFIG.SHEETS.STOCK);
  var rowIndex = findRowIndexById_(sheet, 'ItemID', itemId);
  if (rowIndex < 0) return;
  var rows = getRows_(sheet);
  var row = rows.filter(function (r) { return r._row === rowIndex; })[0];
  var newQty = Math.max(0, (Number(row.Qty) || 0) + delta);
  updateRowByIndex_(sheet, CONFIG.HEADERS.Stock, rowIndex, Object.assign({}, row, { Qty: newQty, UpdatedAt: nowStr_() }));
}

function deleteStockItem(token, id) {
  return safeCall_(function () {
    var session = requireAuth_(token, 'admin');
    var sheet = getSheet_(CONFIG.SHEETS.STOCK);
    var rowIndex = findRowIndexById_(sheet, 'ItemID', id);
    if (rowIndex < 0) throw ValidationError_('ไม่พบสินค้า');
    deleteRowByIndex_(sheet, rowIndex);
    writeAudit_(session, 'DELETE', 'Stock', id, '');
    invalidateStockLiteCache_();
    return { ok: true };
  });
}

/**
 * Adjust one item's stock. payload: {itemId, date, kind, qty, note} where
 * kind is 'in' (รับเข้า: +qty), 'out' (เบิกออก: -qty) or 'set' (ปรับยอด:
 * the counted remaining becomes qty). The starting point is the item's
 * current Stock.Qty (= its latest StockLog QtyRemaining).
 */
function adjustStock(token, payload) {
  return safeCall_(function () {
    var session = requireAuth_(token, 'staff');
    payload = payload || {};
    requireFields_(payload, { itemId: 'สินค้า', date: 'วันที่', kind: 'ประเภทการปรับ' });
    requireDateStr_(payload.date, 'วันที่');
    if (['in', 'out', 'set'].indexOf(payload.kind) === -1) throw ValidationError_('ประเภทการปรับไม่ถูกต้อง');
    var qty = requireNumber_(payload.qty, 'จำนวน', { min: 0 });

    var lock = LockService.getScriptLock();
    lock.waitLock(30000);
    try {
      var item = getRows_(getSheet_(CONFIG.SHEETS.STOCK)).filter(function (r) { return r.ItemID === payload.itemId; })[0];
      if (!item) throw ValidationError_('ไม่พบสินค้า');
      var current = Number(item.Qty) || 0;
      var qtyIn = 0, qtyOut = 0, remaining;
      if (payload.kind === 'in') { qtyIn = qty; remaining = current + qty; }
      else if (payload.kind === 'out') {
        if (qty > current) throw ValidationError_('จำนวนเบิกออก (' + qty + ') มากกว่าจำนวนคงเหลือ (' + current + ')');
        qtyOut = qty; remaining = current - qty;
      } else {
        remaining = qty;
        if (qty > current) qtyIn = qty - current; else qtyOut = current - qty;
      }
      var logId = 'SL' + nextSeq_('STOCKLOG', 5);
      appendRow_(getSheet_(CONFIG.SHEETS.STOCK_LOG), CONFIG.HEADERS.StockLog,
        stockLogRow_(session, logId, payload.date, item, payload.kind, qtyIn, qtyOut, remaining, payload.note));
      resyncStockQtyFromLog_(item.ItemID);
      writeAudit_(session, 'ADJUST', 'StockLog', logId, item.Name + ' ' + payload.kind + ' ' + qty);
      invalidateStockLiteCache_();
      return { ok: true, data: { logId: logId, qtyRemaining: remaining } };
    } finally {
      lock.releaseLock();
    }
  });
}

/** Reserves `n` consecutive sequence numbers in one Counters read/write (bulk imports). */
function nextSeqBlock_(counterKey, n, pad) {
  if (n <= 0) return [];
  var lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    var sheet = getSheet_(CONFIG.SHEETS.COUNTERS);
    var found = getRows_(sheet).filter(function (r) { return r.Key === counterKey; })[0];
    var last = found ? Number(found.LastSeq) || 0 : 0;
    if (found) updateRowByIndex_(sheet, CONFIG.HEADERS.Counters, found._row, { Key: counterKey, LastSeq: last + n });
    else appendRow_(sheet, CONFIG.HEADERS.Counters, { Key: counterKey, LastSeq: n });
    var out = [];
    for (var i = 1; i <= n; i++) out.push(String(last + i).padStart(pad || 4, '0'));
    return out;
  } finally {
    lock.releaseLock();
  }
}

/**
 * CSV upload. payload: {type, rows:[{name, category, unit, minQty, qty,
 * costPrice, prices:[6]}]} — already parsed client-side. Rows are matched to
 * existing items of the same type by name (trimmed, case-insensitive):
 * matches are updated, the rest created. A blank cell leaves that field as
 * it was; a non-blank จำนวนคงเหลือ that differs from the current quantity is
 * written as an 'import' StockLog entry dated today. ราคาทุน is ignored for
 * non-admins. Unknown หมวดหมู่ names are added as stock categories.
 * Writes are batched (one setValues per sheet) so a few hundred rows fit
 * comfortably inside the Apps Script time limit.
 */
function importStockCsv(token, payload) {
  return safeCall_(function () {
    var session = requireAuth_(token, 'staff');
    payload = payload || {};
    var type = payload.type === 'medicine' ? 'medicine' : 'supply';
    var input = Array.isArray(payload.rows) ? payload.rows : [];
    if (!input.length) throw ValidationError_('ไม่พบรายการในไฟล์ CSV');
    var isAdmin = session.level === 'admin';

    // Validate everything before writing anything.
    var clean = input.map(function (r, i) {
      var line = ' (แถวที่ ' + (i + 1) + ')';
      var name = String(r.name || '').trim();
      if (!name) throw ValidationError_('ไม่มีชื่อรายการ' + line);
      var prices = Array.isArray(r.prices) ? r.prices : [];
      return {
        name: name, category: String(r.category || '').trim(), unit: String(r.unit || '').trim(),
        minQty: optionalNumber_(r.minQty, 'จำนวนสต๊อกขั้นต่ำ' + line),
        qty: optionalNumber_(r.qty, 'จำนวนคงเหลือ' + line),
        costPrice: isAdmin ? optionalNumber_(r.costPrice, 'ราคาทุน' + line) : '',
        prices: STOCK_PRICE_TIERS.map(function (t, k) { return optionalNumber_(prices[k], 'ราคาขาย ' + t + ' ชิ้น' + line); })
      };
    });

    var lock = LockService.getScriptLock();
    lock.waitLock(30000);
    try {
      var stockSheet = getSheet_(CONFIG.SHEETS.STOCK);
      var headers = CONFIG.HEADERS.Stock;
      var allRows = getRows_(stockSheet);
      var byName = {};
      allRows.forEach(function (r) { if (stockTypeOf_(r) === type) byName[String(r.Name).trim().toLowerCase()] = r; });

      var newCount = clean.filter(function (r, i) {
        var key = r.name.toLowerCase();
        // a name repeated inside the file only creates one item
        return !byName[key] && clean.map(function (x) { return x.name.toLowerCase(); }).indexOf(key) === i;
      }).length;
      var newIds = nextSeqBlock_('STOCK', newCount, 4);
      var created = [], updated = 0, logs = [], now = nowStr_(), today = todayStr_();

      clean.forEach(function (r) {
        var key = r.name.toLowerCase();
        var row = byName[key];
        var isNew = !row;
        if (isNew) {
          row = { ItemID: 'ST' + newIds.shift(), Name: r.name, Category: '', Unit: '', Qty: 0, CreatedAt: now, StockType: type };
          byName[key] = row;
          created.push(row);
        } else if (created.indexOf(row) === -1) {
          updated++;
        }
        if (r.category) row.Category = r.category;
        if (r.unit) row.Unit = r.unit;
        if (r.minQty !== '') row.MinQty = r.minQty;
        if (r.costPrice !== '') row.CostPrice = r.costPrice;
        STOCK_PRICE_TIERS.forEach(function (t, k) { if (r.prices[k] !== '') row['Price' + t] = r.prices[k]; });
        row.UpdatedAt = now;
        if (r.qty !== '') {
          var current = Number(row.Qty) || 0;
          if (isNew || r.qty !== current) {
            logs.push({ row: row, qtyIn: Math.max(0, r.qty - current), qtyOut: Math.max(0, current - r.qty), remaining: r.qty });
            row.Qty = r.qty;
          }
        }
      });

      // One rewrite of the existing block (blank rows kept as they are) + one append block for new rows.
      var toValues = function (obj) { return headers.map(function (h) { return obj[h] !== undefined ? obj[h] : ''; }); };
      if (allRows.length) {
        var lastRow = stockSheet.getLastRow();
        var block = stockSheet.getRange(2, 1, lastRow - 1, headers.length).getValues();
        allRows.forEach(function (r) { block[r._row - 2] = toValues(r); });
        stockSheet.getRange(2, 1, lastRow - 1, headers.length).setValues(block);
      }
      if (created.length) {
        stockSheet.getRange(stockSheet.getLastRow() + 1, 1, created.length, headers.length).setValues(created.map(toValues));
      }
      if (logs.length) {
        var logIds = nextSeqBlock_('STOCKLOG', logs.length, 5);
        var logSheet = getSheet_(CONFIG.SHEETS.STOCK_LOG);
        var logValues = logs.map(function (l, i) {
          var obj = stockLogRow_(session, 'SL' + logIds[i], today, l.row, 'import', l.qtyIn, l.qtyOut, l.remaining, 'นำเข้า CSV');
          return CONFIG.HEADERS.StockLog.map(function (h) { return obj[h] !== undefined ? obj[h] : ''; });
        });
        logSheet.getRange(logSheet.getLastRow() + 1, 1, logValues.length, CONFIG.HEADERS.StockLog.length).setValues(logValues);
      }

      // Add any หมวดหมู่ the file used that isn't a stock category yet.
      var catSheet = getSheet_(CONFIG.SHEETS.CATEGORIES);
      var knownCats = {};
      getRows_(catSheet).forEach(function (c) { if (c.Type === 'stock') knownCats[c.Name] = true; });
      var newCats = [];
      clean.forEach(function (r) { if (r.category && !knownCats[r.category]) { knownCats[r.category] = true; newCats.push(r.category); } });
      newCats.forEach(function (name) {
        appendRow_(catSheet, CONFIG.HEADERS.Categories, { CategoryID: 'C' + nextSeq_('CATEGORY', 3), Type: 'stock', Name: name, SortOrder: 0 });
      });
      if (newCats.length) invalidateMasterCache_();

      writeAudit_(session, 'IMPORT', 'Stock', '', type + ': +' + created.length + ' / ~' + updated + ' / log ' + logs.length);
      invalidateStockLiteCache_();
      return { ok: true, data: { created: created.length, updated: updated, logged: logs.length, newCategories: newCats } };
    } finally {
      lock.releaseLock();
    }
  });
}

/* ---- อัพเดตสต๊อก (round 25): StockLog — repeatable per-item movement
   history, batch-saved (one Date + one Note shared across N item lines per
   save), modeled directly on the Orders/OrderLines header+lines pattern
   above. See the StockLog HEADERS comment near the top of this file for the
   full schema rationale, and latestStockLogByItem_/resyncStockQtyFromLog_
   above for how Stock.Qty is kept in sync as a denormalized cache. ---- */

/** payload: {date, note, items:[{itemId, qtyIn, qtyRemaining}]} — writes one StockLog row per item (all sharing date/note), then re-syncs Stock.Qty for every touched item. */
function saveStockUpdate(token, payload) {
  return safeCall_(function () {
    var session = requireAuth_(token, 'staff');
    payload = payload || {};
    requireFields_(payload, { date: 'วันที่' });
    requireDateStr_(payload.date, 'วันที่');
    var items = Array.isArray(payload.items) ? payload.items : [];
    if (!items.length) throw ValidationError_('กรุณาเพิ่มรายการสินค้าอย่างน้อย 1 รายการ');

    var stockSheet = getSheet_(CONFIG.SHEETS.STOCK);
    var stockById = {};
    getRows_(stockSheet).forEach(function (r) { stockById[r.ItemID] = r; });
    items.forEach(function (it, i) {
      requireFields_(it, { itemId: 'ชื่อสินค้า (บรรทัดที่ ' + (i + 1) + ')' });
      if (!stockById[it.itemId]) throw ValidationError_('ไม่พบสินค้า (บรรทัดที่ ' + (i + 1) + ')');
      requireNumber_(it.qtyIn || 0, 'จำนวนรับเข้า (บรรทัดที่ ' + (i + 1) + ')', { min: 0 });
      requireNumber_(it.qtyRemaining, 'จำนวนคงเหลือ (บรรทัดที่ ' + (i + 1) + ')', { min: 0 });
    });

    var lock = LockService.getScriptLock();
    lock.waitLock(30000);
    try {
      var logSheet = getSheet_(CONFIG.SHEETS.STOCK_LOG);
      var createdIds = [];
      items.forEach(function (it) {
        var logId = 'SL' + nextSeq_('STOCKLOG', 5);
        var item = stockById[it.itemId];
        var qtyIn = Number(it.qtyIn) || 0, remaining = Number(it.qtyRemaining);
        // จำนวนออก = whatever left between the previous count (+ this receipt) and the new count.
        var qtyOut = Math.max(0, (Number(item.Qty) || 0) + qtyIn - remaining);
        appendRow_(logSheet, CONFIG.HEADERS.StockLog, stockLogRow_(session, logId, payload.date, item, 'update', qtyIn, qtyOut, remaining, payload.note));
        item.Qty = remaining; // same item twice in one batch chains off the first line
        createdIds.push(logId);
      });
      var touchedIds = {};
      items.forEach(function (it) { touchedIds[it.itemId] = true; });
      Object.keys(touchedIds).forEach(resyncStockQtyFromLog_);
      writeAudit_(session, 'CREATE', 'StockLog', createdIds.join(','), payload.date + ' x' + items.length + ' รายการ');
      invalidateStockLiteCache_();
      return { ok: true, data: { logIds: createdIds } };
    } finally {
      lock.releaseLock();
    }
  });
}

/** Full movement history for one item, newest first — shown in the Stock page's history side panel. Rows from before QtyOut existed get it derived from the previous row's remaining (prev + in - remaining). */
function getStockLog(token, itemId) {
  return safeCall_(function () {
    requireAuth_(token, 'staff');
    var rows = getRows_(getSheet_(CONFIG.SHEETS.STOCK_LOG)).filter(function (l) { return l.ItemID === itemId; });
    rows.sort(function (a, b) {
      var d = dateStr_(a.Date).localeCompare(dateStr_(b.Date));
      if (d !== 0) return d;
      return String(a.CreatedAt).localeCompare(String(b.CreatedAt));
    });
    var prev = 0;
    var out = rows.map(function (l) {
      var m = mapStockLogOut_(l);
      if (l.QtyOut === '' || l.QtyOut === null || l.QtyOut === undefined) m.qtyOut = Math.max(0, prev + m.qtyIn - m.qtyRemaining);
      prev = m.qtyRemaining;
      return m;
    });
    return { ok: true, data: out.reverse() };
  });
}

function mapStockLogOut_(l) {
  return {
    id: l.LogID, date: dateStr_(l.Date), itemId: l.ItemID, itemName: l.ItemName,
    qtyIn: Number(l.QtyIn) || 0, qtyOut: Number(l.QtyOut) || 0, qtyRemaining: Number(l.QtyRemaining) || 0,
    note: l.Note, kind: STOCK_LOG_KINDS[l.Kind] ? l.Kind : 'update', createdBy: l.CreatedBy || '', row: l._row
  };
}

/** entry: {id, date, qtyIn, qtyOut, qtyRemaining, note} — edits ONE StockLog row (e.g. fixing a typo in a single item's line after the fact); does not touch the other item rows from that same original batch save. */
function editStockLogEntry(token, entry) {
  return safeCall_(function () {
    var session = requireAuth_(token, 'staff');
    requireFields_(entry, { id: 'รายการ', date: 'วันที่' });
    requireDateStr_(entry.date, 'วันที่');
    requireNumber_(entry.qtyIn || 0, 'จำนวนเข้า', { min: 0 });
    requireNumber_(entry.qtyOut || 0, 'จำนวนออก', { min: 0 });
    requireNumber_(entry.qtyRemaining, 'จำนวนคงเหลือ', { min: 0 });

    var lock = LockService.getScriptLock();
    lock.waitLock(30000);
    try {
      var logSheet = getSheet_(CONFIG.SHEETS.STOCK_LOG);
      var rowIndex = findRowIndexById_(logSheet, 'LogID', entry.id);
      if (rowIndex < 0) throw ValidationError_('ไม่พบรายการ');
      var existing = getRows_(logSheet).filter(function (l) { return l._row === rowIndex; })[0];
      var obj = Object.assign({}, existing, {
        Date: entry.date, QtyIn: Number(entry.qtyIn) || 0, QtyOut: Number(entry.qtyOut) || 0,
        QtyRemaining: Number(entry.qtyRemaining), Note: entry.note || ''
      });
      updateRowByIndex_(logSheet, CONFIG.HEADERS.StockLog, rowIndex, obj);
      resyncStockQtyFromLog_(existing.ItemID);
      writeAudit_(session, 'UPDATE', 'StockLog', entry.id, obj.ItemName);
      invalidateStockLiteCache_();
      return { ok: true, data: mapStockLogOut_(obj) };
    } finally {
      lock.releaseLock();
    }
  });
}

function deleteStockLogEntry(token, id) {
  return safeCall_(function () {
    var session = requireAuth_(token, 'admin');
    var logSheet = getSheet_(CONFIG.SHEETS.STOCK_LOG);
    var rowIndex = findRowIndexById_(logSheet, 'LogID', id);
    if (rowIndex < 0) throw ValidationError_('ไม่พบรายการ');
    var existing = getRows_(logSheet).filter(function (l) { return l._row === rowIndex; })[0];
    deleteRowByIndex_(logSheet, rowIndex);
    resyncStockQtyFromLog_(existing.ItemID); // falls through to a no-op if that was the item's last remaining entry
    writeAudit_(session, 'DELETE', 'StockLog', id, existing.ItemName);
    invalidateStockLiteCache_();
    return { ok: true };
  });
}

/* ======================================================================
   บันทึกข้อมูลการรักษา + นัดหมาย + ออกใบเสร็จ — Records
   Status enum: 'unpaid' | 'paid' | 'cancelled'
   SuppliesJSON: JSON array of {itemId,name,qty,unit} — "รายการเวชภัณฑ์ที่ใช้",
     gated by UseSupplies; names are suggested from เวชภัณฑ์-type Stock items
     and lines with an itemId are taken out of Stock (see recordEffects_).
   PromoJSON: JSON array of {promoId,name,price,doctorSharePct,amount} — round
     13 "รายการโปร"; doctorSharePct is snapshotted server-side from the
     Promotions sheet at save time (never trust the client's copy).
   MedicineJSON: "รายการเบิกยาหน้าคลีนิค" — JSON array of {itemId, name,
     qtyDispensed, qtyUsed, qtyDeposit, unitPrice, amount}: qtyDispensed is
     taken out of Stock, qtyDeposit (= dispensed - used) goes into the
     patient's MedDeposits balance, amount = qtyDispensed * unitPrice is
     charged. Lines saved before this shape existed are {name, amount}.
   OwnMedJSON {name,qty} "รายการถือยาเข้า", DepositUseJSON {itemId,name,qty}
     "รายการใช้ยาฝาก" (taken out of MedDeposits), MessengerJSON {name,qty,
     note} "รายการ Messenger/พัสดุ" — none of these three are charged.
   StockApplied: 'TRUE' on records created since stock/deposit effects
     existed — only those records' lines move Stock/MedDeposits (see
     recordEffects_), so editing an older record never re-deducts it.
   OtherItemsJSON: JSON array of {label,amount} — round 13 itemized
     "รายการอื่นๆ" (OtherAmount column keeps the denormalized sum).
   ====================================================================== */

function parseSupplies_(json) {
  if (!json) return [];
  try { var arr = JSON.parse(json); return Array.isArray(arr) ? arr : []; } catch (e) { return []; }
}
// Round 13: PromoJSON/MedicineJSON/OtherItemsJSON share the same "JSON array
// or blank" shape as SuppliesJSON — reuse the same lenient parser.
var parseLineItems_ = parseSupplies_;

function mapRecordOut_(r) {
  var useCanula = String(r.UseCanula) === 'TRUE' || r.UseCanula === true;
  var canulaAmount = Number(r.CanulaAmount) || 0;
  var amount = Number(r.Amount) || 0;
  var otherAmount = Number(r.OtherAmount) || 0;
  var promoAmount = Number(r.PromoAmount) || 0;
  var medicineAmount = Number(r.MedicineAmount) || 0;
  var doctorShareRaw = Number(r.DoctorShareRaw) || 0;
  var salesShare = Number(r.SalesShare) || 0;
  // grossBase mirrors computeRecordShares_/recomputeDoctorDayShares_: the
  // pool that ส่วนแบ่งแพทย์/เซลล์/คลินิก are carved out of is the procedure
  // amount + promo amount + canula (see the round-13 formula comment below).
  var grossBase = round2_(amount + promoAmount + (useCanula ? canulaAmount : 0));
  return {
    id: r.RID, date: dateStr_(r.Date), patientId: r.PatientID, patientName: r.PatientName, phone: r.Phone,
    serviceCategory: r.ServiceCategory, detail: r.Detail, doctorId: r.DoctorID, doctorName: r.DoctorName,
    salesId: r.SalesID, salesName: r.SalesName, paymentMethod: r.PaymentMethod,
    supplies: parseSupplies_(r.SuppliesJSON), useCanula: useCanula,
    // Records saved before UseSupplies existed have a blank flag — treat any
    // supplies they already carry as checked so editing them keeps the list.
    useSupplies: String(r.UseSupplies) === 'TRUE' || r.UseSupplies === true ||
      (r.UseSupplies === '' && parseSupplies_(r.SuppliesJSON).length > 0),
    canulaAmount: canulaAmount, amount: amount, status: r.Status || 'unpaid',
    receiptNo: r.ReceiptNo || '', note: r.Note || '', createdBy: r.CreatedBy, row: r._row,
    // Revenue-split fields (see computeRecordShares_/recomputeDoctorDayShares_).
    // doctorShareRaw/clinicShareRaw are the plain per-record commission split
    // BEFORE the doctor's daily-minimum-guarantee top-up is applied — the
    // per-record detail panel shows these (not doctorShare/clinicShare, which
    // are the topped-up, actually-paid-out amounts) since the guarantee only
    // makes sense aggregated across a whole day; see computeDailySummaryDetail_.
    otherAmount: otherAmount,
    salesCommissionPct: Number(r.SalesCommissionPct) || 0,
    salesShare: salesShare,
    doctorCommissionPct: Number(r.DoctorCommissionPct) || 0,
    doctorShare: Number(r.DoctorShare) || 0,
    doctorShareRaw: doctorShareRaw,
    doctorShareNote: r.DoctorShareNote || '',
    clinicShare: Number(r.ClinicShare) || 0,
    clinicShareRaw: round2_(grossBase - salesShare - doctorShareRaw),
    // Round 13 line-item groups.
    usePromo: String(r.UsePromo) === 'TRUE' || r.UsePromo === true,
    promoItems: parseLineItems_(r.PromoJSON), promoAmount: promoAmount,
    useMedicine: String(r.UseMedicine) === 'TRUE' || r.UseMedicine === true,
    medicineItems: parseLineItems_(r.MedicineJSON), medicineAmount: medicineAmount,
    useOther: String(r.UseOther) === 'TRUE' || r.UseOther === true,
    otherItems: parseLineItems_(r.OtherItemsJSON),
    useOwnMed: String(r.UseOwnMed) === 'TRUE', ownMedItems: parseLineItems_(r.OwnMedJSON),
    useDepositUse: String(r.UseDepositUse) === 'TRUE', depositUseItems: parseLineItems_(r.DepositUseJSON),
    useMessenger: String(r.UseMessenger) === 'TRUE', messengerItems: parseLineItems_(r.MessengerJSON),
    stockApplied: String(r.StockApplied) === 'TRUE',
    supplierPayAmount: Number(r.SupplierPayAmount) || round2_(medicineAmount + otherAmount),
    estimatedRevenue: grossBase,
    totalCollected: round2_(grossBase + medicineAmount + otherAmount),
    visitStatus: r.VisitStatus || 'scheduled',
    // Round 14 item 1.7 — OPD photo attachment (Drive fileId; fetch its
    // viewable data URI via getPhotoUrl, same as patient/login/logo photos).
    opdImageFileId: r.OpdImageFileId || ''
  };
}

/* ======================================================================
   REVENUE SPLIT ENGINE — see the spec's formulas (ยอดชำระจากลูกค้า /
   ส่วนแบ่งเซลล์ / ส่วนแบ่งแพทย์ / ส่วนแบ่งคลินิก) and the confirmed
   clarifications: (1) "สถานะ" (payment: ชำระแล้ว/รอชำระ/ยกเลิก) stays as-is
   and still gates revenue inclusion; VisitStatus (นัดหมาย/เสร็จสิ้น) is a
   separate, independent workflow marker. (2) OtherAmount/MedicineAmount
   ("จำนวนเงินอื่นๆ"/"จำนวนเงินยา") are tracked as their own bucket
   (SupplierPayAmount) and are NOT added into anyone's revenue share.
   ------------------------------------------------------------------------
   ROUND 13 UPDATE — ส่วนแบ่งโปรโมชัน. Per the client's explicit spec this
   round REVERSES round 12's "Promotions is standalone" decision: each
   รายการโปร line's own DoctorSharePct (from the Promotions sheet, snapshotted
   server-side) is now used — NOT the doctor's own CommissionPct — to compute
   that line's contribution to ส่วนแบ่งแพทย์. Concretely, for a record with
   procedure amount A, promo lines summing to P (with a blended
   promo-doctor-share PD = sum(line.amount * line.doctorSharePct/100)), doctor
   commission % Cd, sales commission % Cs, and canula amount K (only if
   useCanula):
     salesShare      = round(A * Cs / 100)                     -- (2) sales never sees promo money
     doctorShareRaw  = round(A * Cd / 100) + PD                -- (1)
     grossBase       = A + P + K                                -- everything the 3-way split is carved from
     clinicShareRaw  = grossBase - salesShare - doctorShareRaw  -- (3) "ที่เหลือ" of both A and P, plus K in full
   With P=0 this is identical to the pre-round-13 formulas, so historical
   records (PromoAmount blank/0) are unaffected.
   ====================================================================== */

function round2_(n) {
  return Math.round((Number(n) || 0) * 100) / 100;
}

/**
 * Raw (pre-daily-minimum) commission split for one record.
 * @param {number} customerAmount    "จำนวนเงินหัตถการ" — procedure amount only.
 * @param {number} promoAmount       Sum of this record's รายการโปร line amounts.
 * @param {number} promoDoctorShare  Sum of (line.amount * line.doctorSharePct/100)
 *   across this record's รายการโปร lines — the caller (saveRecord) computes
 *   this from server-trusted Promotions data, never from client input.
 * @param {number} canulaAmount      Canula amount, already zeroed by the
 *   caller if useCanula is false.
 */
function computeRecordShares_(customerAmount, promoAmount, promoDoctorShare, canulaAmount, doctor, sales) {
  var salesPct = sales ? (Number(sales.CommissionRate) || 0) : 0;
  var doctorPct = doctor ? (Number(doctor.CommissionPct) || 0) : 0;
  var salesShare = round2_((Number(customerAmount) || 0) * salesPct / 100);
  var doctorShareRaw = round2_((Number(customerAmount) || 0) * doctorPct / 100 + (Number(promoDoctorShare) || 0));
  var grossBase = round2_((Number(customerAmount) || 0) + (Number(promoAmount) || 0) + (Number(canulaAmount) || 0));
  return {
    salesPct: salesPct,
    salesShare: salesShare,
    doctorPct: doctorPct,
    doctorShareRaw: doctorShareRaw,
    grossBase: grossBase,
    clinicShareRaw: round2_(grossBase - salesShare - doctorShareRaw)
  };
}

/**
 * Enforces a doctor's daily minimum guarantee across every one of their
 * (non-cancelled) records on one date: if the sum of raw commission-based
 * DoctorShareRaw values for the day is below the doctor's configured
 * MinAmount, the shortfall is topped up entirely onto the most recently
 * saved record in the group (flagged with a DoctorShareNote so it's
 * auditable in one place), and every other record in the group is reset to
 * its plain raw share. ClinicShare is recomputed for every touched record
 * since it depends on DoctorShare. Must be called from inside the caller's
 * LockService lock (saveRecord/deleteRecord) since it rewrites sibling
 * rows on the Records sheet.
 */
// Round 19 item 4: `allRows` is an optional pre-fetched Records getRows_()
// snapshot — when a caller (saveRecord) already has a fresh, POST-write
// snapshot of the whole sheet in hand, passing it in here avoids yet
// another independent full-sheet read (a real cost once Records passes
// 5000+ rows). Callers that don't have one (deleteRecord — the row shift
// from deleteRowByIndex_ means any pre-delete snapshot would have stale
// `_row` numbers) simply omit it and get the original fresh-read behavior.
// Also now RETURNS the (possibly-updated) group array so a caller that just
// wrote a record into this exact doctor+date group can read its final
// DoctorShare/ClinicShare straight off the return value instead of doing a
// second full-sheet re-read+filter (see saveRecord).
function recomputeDoctorDayShares_(doctorId, date, allRows) {
  if (!doctorId || !date) return [];
  var doctorRows = getRows_(getSheet_(CONFIG.SHEETS.DOCTORS));
  var doctor = doctorRows.filter(function (d) { return d.DoctorID === doctorId; })[0];
  var minAmount = doctor ? (Number(doctor.MinAmount) || 0) : 0;

  var sheet = getSheet_(CONFIG.SHEETS.RECORDS);
  var rows = allRows || getRows_(sheet);
  var group = rows.filter(function (r) {
    return r.DoctorID === doctorId && dateStr_(r.Date) === date && r.Status !== 'cancelled';
  });
  if (!group.length) return [];
  group.sort(function (a, b) { return String(a.CreatedAt).localeCompare(String(b.CreatedAt)); });

  var rawTotal = 0;
  group.forEach(function (r) { rawTotal += Number(r.DoctorShareRaw) || 0; });
  var shortfall = minAmount > rawTotal ? round2_(minAmount - rawTotal) : 0;
  var lastIdx = group.length - 1;

  return group.map(function (r, i) {
    var raw = Number(r.DoctorShareRaw) || 0;
    var finalShare = raw;
    var note = '';
    if (shortfall > 0 && i === lastIdx) {
      finalShare = round2_(raw + shortfall);
      note = 'ปรับเพิ่มให้ครบยอดขั้นต่ำต่อวัน (' + minAmount.toLocaleString('en-US') + ' บาท) ของ ' + (doctor ? doctor.Name : doctorId);
    }
    var customerAmount = Number(r.Amount) || 0;
    var promoAmount = Number(r.PromoAmount) || 0;
    var canula = (String(r.UseCanula) === 'TRUE' || r.UseCanula === true) ? (Number(r.CanulaAmount) || 0) : 0;
    var salesShare = Number(r.SalesShare) || 0;
    // grossBase = A + P + K (see the round-13 formula comment above
    // computeRecordShares_) — with PromoAmount 0/blank this is identical to
    // the pre-round-13 "customerAmount + canula" base.
    var grossBase = round2_(customerAmount + promoAmount + canula);
    var clinicShare = round2_(grossBase - salesShare - finalShare);
    var updated = Object.assign({}, r, { DoctorShare: finalShare, DoctorShareNote: note, ClinicShare: clinicShare });
    if (Number(r.DoctorShare) !== finalShare || String(r.DoctorShareNote || '') !== note || Number(r.ClinicShare) !== clinicShare) {
      updateRowByIndex_(sheet, CONFIG.HEADERS.Records, r._row, updated);
    }
    return updated;
  });
}

function listRecords(token, opts) {
  return safeCall_(function () {
    requireAuth_(token, 'staff');
    opts = opts || {};
    var filters = compactFilters_(opts.filters);
    var rows = getRows_(getSheet_(CONFIG.SHEETS.RECORDS));
    if (filters.q) {
      var q = String(filters.q).trim().toLowerCase();
      var qDigits = q.replace(/[^\d]/g, '');
      rows = rows.filter(function (r) {
        return String(r.PatientName).toLowerCase().indexOf(q) > -1 || (qDigits && String(r.Phone).indexOf(qDigits) > -1);
      });
    }
    if (filters.doctorId) rows = rows.filter(function (r) { return r.DoctorID === filters.doctorId; });
    if (filters.status) rows = rows.filter(function (r) { return r.Status === filters.status; });
    if (filters.dateFrom) rows = rows.filter(function (r) { return dateStr_(r.Date) >= filters.dateFrom; });
    if (filters.dateTo) rows = rows.filter(function (r) { return dateStr_(r.Date) <= filters.dateTo; });
    // Sort by date descending (latest first) — compare normalized 'yyyy-MM-dd'
    // strings, not the raw Date objects: String(dateObj) starts with the
    // weekday name ("Sat Aug 16 2026..."), which sorts alphabetically and NOT
    // chronologically, silently scrambling the order (round 12 fix, item 1.1.1).
    rows.sort(function (a, b) {
      return dateStr_(b.Date).localeCompare(dateStr_(a.Date)) || String(b.CreatedAt).localeCompare(String(a.CreatedAt));
    });
    var page = paginate_(rows, opts.page, opts.pageSize);
    page.items = page.items.map(mapRecordOut_);
    return { ok: true, data: page };
  });
}

/** Records grouped by day for a given year/month, for the calendar view toggle. */
function listRecordsForCalendar(token, year, month) {
  return safeCall_(function () {
    requireAuth_(token, 'staff');
    var prefix = year + '-' + ('0' + month).slice(-2);
    var rows = getRows_(getSheet_(CONFIG.SHEETS.RECORDS)).filter(function (r) { return dateStr_(r.Date).indexOf(prefix) === 0; });
    var byDay = {};
    rows.forEach(function (r) {
      var d = dateStr_(r.Date);
      if (!byDay[d]) byDay[d] = [];
      byDay[d].push(mapRecordOut_(r));
    });
    return { ok: true, data: byDay };
  });
}

/**
 * Aggregated daily reconciliation summary for the Records calendar's
 * day-detail side panel, also reused verbatim by the LINE daily push (see
 * sendLineDailySummary) so both surfaces always agree. Only counts records
 * with Status === 'paid' for the given date — unpaid isn't collected yet
 * and cancelled isn't real revenue.
 *
 * Pure aggregation, no auth check — shared by getDailySummaryDetail (the
 * Records-calendar day panel) and sendLineDailySummary (the LINE push),
 * so both surfaces show exactly the same numbers in exactly the same
 * shape. Doctor/sales shares are grouped by person using each record's
 * FINAL post-minimum-guarantee DoctorShare/SalesShare (not itemized per
 * treatment, and not the raw pre-top-up amount) per the clinic's spec.
 */
function computeDailySummaryDetail_(dateStr) {
  var records = getRows_(getSheet_(CONFIG.SHEETS.RECORDS)).filter(function (r) {
    return dateStr_(r.Date) === dateStr && r.Status === 'paid';
  });

  var totalCollected = 0, cashAmount = 0, transferAmount = 0, cardAmount = 0, otherAmountTotal = 0, canulaTotal = 0, clinicNet = 0;
  // Round 13: promoAmountTotal/medicineAmountTotal/supplierPayTotal are new,
  // additive stats — promo money IS included in totalCollected (it's money
  // actually paid by the patient, see mapRecordOut_.totalCollected), while
  // medicine (like OtherAmount before it) is tracked but excluded from
  // clinicNet since it's a supplier cost pass-through, not clinic revenue.
  var promoAmountTotal = 0, medicineAmountTotal = 0;
  var doctorMap = {}, salesMap = {}, otherItems = [], medicineItems = [], promoMap = {};
  records.forEach(function (r) {
    var useCanula = String(r.UseCanula) === 'TRUE' || r.UseCanula === true;
    var canula = useCanula ? (Number(r.CanulaAmount) || 0) : 0;
    var other = Number(r.OtherAmount) || 0;
    var promo = Number(r.PromoAmount) || 0;
    var medicine = Number(r.MedicineAmount) || 0;
    var amount = Number(r.Amount) || 0;
    var collected = round2_(amount + promo + canula + medicine + other);
    totalCollected += collected;
    canulaTotal += canula;
    otherAmountTotal += other;
    promoAmountTotal += promo;
    medicineAmountTotal += medicine;
    clinicNet += Number(r.ClinicShare) || 0;
    if (r.PaymentMethod === 'เงินสด') cashAmount += collected;
    else if (r.PaymentMethod === 'โอนเงิน') transferAmount += collected;
    else if (r.PaymentMethod === 'บัตรเครดิต') cardAmount += collected;
    if (r.DoctorID) {
      if (!doctorMap[r.DoctorID]) doctorMap[r.DoctorID] = { name: r.DoctorName, amount: 0, topped: false };
      doctorMap[r.DoctorID].amount += Number(r.DoctorShare) || 0;
      // DoctorShareNote is only ever set by recomputeDoctorDayShares_ when a
      // record's share was topped up to meet the doctor's daily minimum —
      // surfacing that here lets the summary mark the doctor with a footnote
      // instead of silently showing an inflated number with no explanation.
      if (r.DoctorShareNote) doctorMap[r.DoctorID].topped = true;
    }
    if (r.SalesID) {
      if (!salesMap[r.SalesID]) salesMap[r.SalesID] = { name: r.SalesName, amount: 0 };
      salesMap[r.SalesID].amount += Number(r.SalesShare) || 0;
    }
    // Round 15 item 3: itemize down to each individual line item (its own
    // "ชื่อรายการ" name/label), not one aggregate row per record — the day
    // panel's "จ่าย supplier" list previously only showed the record's
    // patient name with no way to tell which product/medicine was bought.
    if (medicine > 0) {
      parseLineItems_(r.MedicineJSON).forEach(function (line) {
        var amt = Number(line.amount) || 0;
        if (amt > 0) medicineItems.push({ recordId: r.RID, patientName: r.PatientName, name: line.name || 'ยา', amount: round2_(amt) });
      });
    }
    if (other > 0) {
      parseLineItems_(r.OtherItemsJSON).forEach(function (line) {
        var amt = Number(line.amount) || 0;
        if (amt > 0) otherItems.push({ recordId: r.RID, patientName: r.PatientName, name: line.label || 'อื่นๆ', amount: round2_(amt) });
      });
    }
    // Round 18 item 1: itemized per-promotion breakdown ("Pro A (x1)") for
    // the day panel/LINE push, grouped by promo name across all of today's
    // records — same "group + count" shape the dashboard's byPromotion
    // aggregation (getDashboardData) already uses, just also tracking a
    // per-name occurrence count here since the day panel wants "(xN)".
    if (promo > 0) {
      parseLineItems_(r.PromoJSON).forEach(function (line) {
        var amt = Number(line.amount) || 0;
        if (amt <= 0) return;
        var key = line.name || 'ไม่ระบุ';
        if (!promoMap[key]) promoMap[key] = { name: key, qty: 0, amount: 0 };
        promoMap[key].qty += 1;
        promoMap[key].amount += amt;
      });
    }
  });

  var doctorShares = Object.keys(doctorMap).map(function (k) { return { name: doctorMap[k].name, amount: round2_(doctorMap[k].amount), topped: doctorMap[k].topped }; })
    .sort(function (a, b) { return b.amount - a.amount; });
  var salesShares = Object.keys(salesMap).map(function (k) { return { name: salesMap[k].name, amount: round2_(salesMap[k].amount) }; })
    .sort(function (a, b) { return b.amount - a.amount; });
  var doctorTotal = round2_(doctorShares.reduce(function (s, d) { return s + d.amount; }, 0));
  var salesTotal = round2_(salesShares.reduce(function (s, d) { return s + d.amount; }, 0));
  var promoItems = Object.keys(promoMap).map(function (k) { return { name: promoMap[k].name, qty: promoMap[k].qty, amount: round2_(promoMap[k].amount) }; })
    .sort(function (a, b) { return b.amount - a.amount; });

  return {
    date: dateStr, patientCount: records.length,
    totalCollected: round2_(totalCollected), cashAmount: round2_(cashAmount), transferAmount: round2_(transferAmount), cardAmount: round2_(cardAmount),
    canulaTotal: round2_(canulaTotal), otherAmountTotal: round2_(otherAmountTotal), clinicNet: round2_(clinicNet),
    doctorShares: doctorShares, doctorTotal: doctorTotal, hasToppedDoctor: doctorShares.some(function (d) { return d.topped; }),
    salesShares: salesShares, salesTotal: salesTotal, otherItems: otherItems,
    promoAmountTotal: round2_(promoAmountTotal), medicineAmountTotal: round2_(medicineAmountTotal), medicineItems: medicineItems,
    // Round 18 item 1: itemized promo breakdown (name + occurrence count +
    // summed amount) — promoAmountTotal above is kept unchanged for backward
    // compat (it already equals the sum of these).
    promoItems: promoItems,
    supplierPayTotal: round2_(medicineAmountTotal + otherAmountTotal)
  };
}

function getDailySummaryDetail(token, dateStr) {
  return safeCall_(function () {
    requireAuth_(token, 'staff');
    requireDateStr_(dateStr, 'วันที่');
    return { ok: true, data: computeDailySummaryDetail_(dateStr) };
  });
}

function getRecord(token, id) {
  return safeCall_(function () {
    requireAuth_(token, 'staff');
    var rows = getRows_(getSheet_(CONFIG.SHEETS.RECORDS));
    var r = rows.filter(function (x) { return x.RID === id; })[0];
    if (!r) throw ValidationError_('ไม่พบข้อมูลการรักษา');
    var out = mapRecordOut_(r);
    // Nickname isn't denormalized onto Records (only Name/Phone are) — look it
    // up from Patients for display in the record modal/panel ("ชื่อเล่น").
    var patient = getRows_(getSheet_(CONFIG.SHEETS.PATIENTS)).filter(function (p) { return p.ID === r.PatientID; })[0];
    out.patientNickname = patient ? (patient.Nickname || '') : '';
    return { ok: true, data: out };
  });
}

/**
 * Validates+sums a repeatable line-item group's `amount` fields (round-13
 * รายการโปร/ยา/รายการอื่นๆ). Returns { items, total } where `items` is the
 * sanitized array actually persisted (only {..., amount} survive, coerced to
 * numbers) and `total` is their round2_'d sum. `useFlag` false (checkbox off)
 * always yields an empty group regardless of what the client sent, so a
 * stale/leftover draft never sneaks money in while its section is collapsed.
 */
function sumLineItems_(items, useFlag, fieldLabel) {
  if (!useFlag || !Array.isArray(items)) return { items: [], total: 0 };
  var total = 0;
  var out = items.map(function (it) {
    var amt = requireNumber_(it && it.amount || 0, fieldLabel, { min: 0 });
    total += amt;
    return Object.assign({}, it, { amount: amt });
  });
  return { items: out, total: round2_(total) };
}

/** รายการเบิกยาหน้าคลีนิค lines -> {items, total}; see the MedicineJSON note above. */
function medicineLines_(record) {
  var out = { items: [], total: 0 };
  if (!record.useMedicine || !Array.isArray(record.medicineItems)) return out;
  record.medicineItems.forEach(function (it, i) {
    if (!it || (!String(it.name || '').trim() && !it.itemId)) return;
    var line = ' (บรรทัดที่ ' + (i + 1) + ')';
    if (!('qtyDispensed' in it)) { // pre-dispense-table line: name + amount only
      var legacyAmt = requireNumber_(it.amount || 0, 'จำนวนเงินยา' + line, { min: 0 });
      out.items.push({ name: it.name || '', amount: legacyAmt });
      out.total += legacyAmt;
      return;
    }
    if (!it.itemId) throw ValidationError_('กรุณาเลือกยาจากรายการสต๊อกยา' + line);
    var qd = requireNumber_(it.qtyDispensed || 0, 'จำนวนที่เบิก' + line, { min: 0 });
    var qu = requireNumber_(it.qtyUsed || 0, 'จำนวนที่ใช้' + line, { min: 0 });
    if (qu > qd) throw ValidationError_('จำนวนที่ใช้มากกว่าจำนวนที่เบิก' + line);
    var up = requireNumber_(it.unitPrice || 0, 'ราคาต่อหน่วย' + line, { min: 0 });
    var amt = round2_(qd * up);
    out.items.push({ itemId: it.itemId, name: String(it.name || '').trim(), qtyDispensed: qd, qtyUsed: qu, qtyDeposit: round2_(qd - qu), unitPrice: up, amount: amt });
    out.total += amt;
  });
  out.total = round2_(out.total);
  return out;
}

/** Uncharged quantity lines (ถือยาเข้า / ใช้ยาฝาก / Messenger): blank names dropped, qty validated, only `fields` kept. */
function qtyLines_(items, use, label, fields, requireItem) {
  if (!use || !Array.isArray(items)) return [];
  return items.filter(function (it) { return it && (String(it.name || '').trim() || it.itemId); }).map(function (it, i) {
    var line = ' (' + label + ' บรรทัดที่ ' + (i + 1) + ')';
    if (requireItem && !it.itemId) throw ValidationError_('กรุณาเลือกรายการ' + line);
    var o = { name: String(it.name || '').trim(), qty: requireNumber_(it.qty || 0, 'จำนวน' + line, { min: 0 }) };
    fields.forEach(function (f) { o[f] = it[f] || ''; });
    return o;
  });
}

/**
 * What a saved Records row takes out of Stock and puts into / takes out of
 * MedDeposits: {stock: {itemId: qty}, deposit: {'patientId|itemId': qty}}.
 * Rows without StockApplied=TRUE (saved before this existed) have none.
 */
function recordEffects_(r) {
  var eff = { stock: {}, deposit: {} };
  if (String(r.StockApplied) !== 'TRUE') return eff;
  var add = function (map, key, n) { if (key && n) map[key] = round2_((map[key] || 0) + n); };
  if (String(r.UseSupplies) === 'TRUE') {
    parseLineItems_(r.SuppliesJSON).forEach(function (l) { add(eff.stock, l.itemId, Number(l.qty) || 0); });
  }
  if (String(r.UseMedicine) === 'TRUE') {
    parseLineItems_(r.MedicineJSON).forEach(function (l) {
      add(eff.stock, l.itemId, Number(l.qtyDispensed) || 0);
      add(eff.deposit, l.itemId && (r.PatientID + '|' + l.itemId), Number(l.qtyDeposit) || 0);
    });
  }
  if (String(r.UseDepositUse) === 'TRUE') {
    parseLineItems_(r.DepositUseJSON).forEach(function (l) { add(eff.deposit, l.itemId && (r.PatientID + '|' + l.itemId), -(Number(l.qty) || 0)); });
  }
  return eff;
}

/**
 * Moves Stock and MedDeposits by (newEff - oldEff) — so a new record applies
 * its lines, an edit applies only what changed, and a delete (newEff empty)
 * puts everything back. Checks every item first and throws before writing
 * anything if stock or a deposit balance would go below zero. Stock moves
 * are StockLog 'record' rows dated today (not the record's date, so they
 * are always the item's latest entry and Stock.Qty follows them).
 * Must run inside the caller's script lock.
 */
function applyRecordEffects_(session, oldEff, newEff, label, patientName) {
  var diff = function (a, b) {
    var out = {};
    Object.keys(a).concat(Object.keys(b)).forEach(function (k) { var d = round2_((b[k] || 0) - (a[k] || 0)); if (d) out[k] = d; });
    return out;
  };
  var stockDelta = diff(oldEff.stock, newEff.stock);     // + = more taken out of stock
  var depositDelta = diff(oldEff.deposit, newEff.deposit); // + = more deposited
  if (!Object.keys(stockDelta).length && !Object.keys(depositDelta).length) return;

  var stockById = {};
  getRows_(getSheet_(CONFIG.SHEETS.STOCK)).forEach(function (r) { stockById[r.ItemID] = r; });
  Object.keys(stockDelta).forEach(function (id) {
    var item = stockById[id];
    if (!item) { if (stockDelta[id] > 0) throw ValidationError_('ไม่พบรายการในสต๊อก (อาจถูกลบไปแล้ว)'); return; }
    var cur = Number(item.Qty) || 0;
    if (stockDelta[id] > cur) throw ValidationError_('สต๊อก ' + item.Name + ' ไม่พอ (คงเหลือ ' + cur + ' ' + item.Unit + ')');
  });
  var depSheet = getSheet_(CONFIG.SHEETS.MED_DEPOSITS);
  var depByKey = {};
  getRows_(depSheet).forEach(function (r) { depByKey[r.PatientID + '|' + r.ItemID] = r; });
  Object.keys(depositDelta).forEach(function (k) {
    var row = depByKey[k], bal = row ? Number(row.QtyRemaining) || 0 : 0;
    if (bal + depositDelta[k] < 0) {
      var name = row ? row.ItemName : (stockById[k.split('|')[1]] || {}).Name || '';
      throw ValidationError_('ยาฝาก ' + name + ' ไม่พอ (คงเหลือ ' + bal + ')');
    }
  });

  var logSheet = getSheet_(CONFIG.SHEETS.STOCK_LOG);
  var movedIds = Object.keys(stockDelta).filter(function (id) { return stockById[id]; });
  movedIds.forEach(function (id) {
    var item = stockById[id], d = stockDelta[id], cur = Number(item.Qty) || 0;
    appendRow_(logSheet, CONFIG.HEADERS.StockLog,
      stockLogRow_(session, 'SL' + nextSeq_('STOCKLOG', 5), todayStr_(), item, 'record', d < 0 ? -d : 0, d > 0 ? d : 0, round2_(cur - d), label));
    resyncStockQtyFromLog_(id);
  });
  if (movedIds.length) invalidateStockLiteCache_();

  Object.keys(depositDelta).forEach(function (k) {
    var row = depByKey[k], parts = k.split('|');
    if (row) {
      updateRowByIndex_(depSheet, CONFIG.HEADERS.MedDeposits, row._row, Object.assign({}, row, {
        Date: todayStr_(), QtyRemaining: round2_((Number(row.QtyRemaining) || 0) + depositDelta[k]), CreatedBy: session.username, UpdatedAt: nowStr_()
      }));
    } else {
      appendRow_(depSheet, CONFIG.HEADERS.MedDeposits, {
        DepositID: 'DP' + nextSeq_('DEPOSIT', 5), Date: todayStr_(), PatientID: parts[0], PatientName: patientName || '',
        ItemName: (stockById[parts[1]] || {}).Name || '', QtyRemaining: depositDelta[k], CreatedBy: session.username, ItemID: parts[1], UpdatedAt: nowStr_()
      });
    }
  });
}

/** ยาฝาก the patient still has (balance > 0), for the record modal. */
function listPatientDeposits(token, patientId) {
  return safeCall_(function () {
    requireAuth_(token, 'staff');
    var rows = getRows_(getSheet_(CONFIG.SHEETS.MED_DEPOSITS)).filter(function (r) { return r.PatientID === patientId && (Number(r.QtyRemaining) || 0) > 0; });
    return { ok: true, data: rows.map(function (r) {
      return { id: r.DepositID, itemId: r.ItemID, name: r.ItemName, qty: Number(r.QtyRemaining) || 0, date: dateStr_(r.Date), createdBy: r.CreatedBy };
    }) };
  });
}

function saveRecord(token, record) {
  return safeCall_(function () {
    var session = requireAuth_(token, 'staff');
    // Round 14 item 1.1: หมวดหมู่การรักษา field removed from the modal UI —
    // no longer a required field. ServiceCategory column stays in the
    // schema (append-only discipline) and still stores whatever the client
    // sends (now always '' going forward); old records keep their value.
    requireFields_(record, {
      date: 'วันที่', patientId: 'คนไข้',
      doctorId: 'แพทย์', paymentMethod: 'วิธีชำระเงิน'
    });
    requireDateStr_(record.date, 'วันที่');
    // "amount" here is จำนวนเงินหัตถการ — the base procedure price that
    // sales/doctor commission % are computed against. Canula/promo/medicine/
    // other money are tracked separately (see header schema comment) — only
    // amount (always) and promo (round 13) feed the commissionable base.
    var amount = requireNumber_(record.amount || 0, 'จำนวนเงินหัตถการ', { min: 0 });
    var canulaAmount = requireNumber_(record.canulaAmount || 0, 'ค่าใช้ Canula', { min: 0 });
    // รายการเวชภัณฑ์ที่ใช้ — quantity only (no money); blank-name lines dropped;
    // lines picked from stock (itemId) are deducted via applyRecordEffects_.
    var supplies = (record.useSupplies && Array.isArray(record.supplies) ? record.supplies : [])
      .filter(function (it) { return it && String(it.name || '').trim(); })
      .map(function (it) {
        return {
          itemId: it.itemId || '', name: String(it.name).trim(),
          qty: requireNumber_(it.qty || 0, 'จำนวนเวชภัณฑ์', { min: 0 }), unit: it.unit || ''
        };
      });

    // Round 13 repeatable line-item groups — sums are ALWAYS recomputed
    // server-side from the line arrays (never trust a client-sent aggregate).
    var medicineGroup = medicineLines_(record);
    var ownMedItems = qtyLines_(record.ownMedItems, !!record.useOwnMed, 'ถือยาเข้า', [], false);
    var depositUseItems = qtyLines_(record.depositUseItems, !!record.useDepositUse, 'ใช้ยาฝาก', ['itemId'], true);
    var messengerItems = qtyLines_(record.messengerItems, !!record.useMessenger, 'Messenger/พัสดุ', ['note'], false);
    var otherGroup = sumLineItems_(record.otherItems, !!record.useOther, 'จำนวนเงินอื่นๆ');
    var otherAmount = otherGroup.total;

    var lock = LockService.getScriptLock();
    lock.waitLock(30000);
    try {
      var sheet = getSheet_(CONFIG.SHEETS.RECORDS);
      var isNew = !record.id;
      var patientRows = getRows_(getSheet_(CONFIG.SHEETS.PATIENTS));
      var patient = patientRows.filter(function (p) { return p.ID === record.patientId; })[0];
      if (!patient) throw ValidationError_('ไม่พบข้อมูลคนไข้ที่เลือก');

      var doctorRows = getRows_(getSheet_(CONFIG.SHEETS.DOCTORS));
      var doctor = doctorRows.filter(function (d) { return d.DoctorID === record.doctorId; })[0];
      if (!doctor) throw ValidationError_('ไม่พบข้อมูลแพทย์ที่เลือก');

      var salesName = '', sales = null;
      if (record.salesId) {
        var salesRows = getRows_(getSheet_(CONFIG.SHEETS.SALES));
        sales = salesRows.filter(function (s) { return s.SalesID === record.salesId; })[0];
        salesName = sales ? sales.Name : '';
      }

      // รายการโปร — each line's price/name/% is re-resolved from the
      // Promotions sheet by promoId (not trusted from the client) so a
      // tampered or stale `doctorSharePct` in the browser can never change
      // the doctor's commission; only the line's `amount` (จำนวนเงินโปร,
      // pre-filled from the promo's price but editable) is taken as-is, same
      // trust level as the treatment amount itself.
      var promoAmount = 0, promoDoctorShare = 0, promoOut = [];
      if (record.usePromo && Array.isArray(record.promoItems) && record.promoItems.length) {
        var allPromos = getRows_(getSheet_(CONFIG.SHEETS.PROMOTIONS));
        record.promoItems.forEach(function (line) {
          var promo = allPromos.filter(function (p) { return p.PromoID === line.promoId; })[0];
          if (!promo) throw ValidationError_('ไม่พบโปรโมชันที่เลือก (อาจถูกลบไปแล้ว)');
          var lineAmount = requireNumber_(line.amount || 0, 'จำนวนเงินโปร', { min: 0 });
          var doctorPct = Number(promo.DoctorSharePct) || 0;
          promoAmount += lineAmount;
          promoDoctorShare += lineAmount * doctorPct / 100;
          promoOut.push({ promoId: promo.PromoID, name: promo.Name, price: Number(promo.Price) || 0, doctorSharePct: doctorPct, amount: lineAmount });
        });
      }
      promoAmount = round2_(promoAmount);
      promoDoctorShare = round2_(promoDoctorShare);

      var oldDoctorId = '', oldDate = '';
      if (!isNew) {
        var existingRows = getRows_(sheet);
        var existing = existingRows.filter(function (r) { return r.RID === record.id; })[0];
        if (!existing) throw ValidationError_('ไม่พบข้อมูลการรักษาที่จะแก้ไข');
        oldDoctorId = existing.DoctorID;
        oldDate = dateStr_(existing.Date);
      }

      var shares = computeRecordShares_(amount, promoAmount, promoDoctorShare, record.useCanula ? canulaAmount : 0, doctor, sales);
      var supplierPayAmount = round2_(medicineGroup.total + otherAmount);

      var obj = {
        RID: isNew ? ('R' + nextSeq_('RECORD', 6)) : record.id,
        Date: record.date, PatientID: patient.ID, PatientName: patient.FirstName + ' ' + patient.LastName,
        Phone: patient.Phone, ServiceCategory: record.serviceCategory, Detail: record.detail || '',
        DoctorID: doctor.DoctorID, DoctorName: doctor.Name, SalesID: record.salesId || '', SalesName: salesName,
        PaymentMethod: record.paymentMethod, SuppliesJSON: JSON.stringify(supplies),
        UseCanula: record.useCanula ? 'TRUE' : 'FALSE', CanulaAmount: record.useCanula ? canulaAmount : 0,
        Amount: amount, Status: record.status || 'unpaid', ReceiptNo: isNew ? '' : (record.receiptNo || ''),
        Note: record.note || '', CreatedBy: isNew ? session.username : (record.createdBy || session.username),
        CreatedAt: isNew ? nowStr_() : (record.createdAt || nowStr_()), UpdatedAt: nowStr_(),
        OtherAmount: otherAmount,
        SalesCommissionPct: shares.salesPct, SalesShare: shares.salesShare,
        DoctorCommissionPct: shares.doctorPct, DoctorShareRaw: shares.doctorShareRaw,
        DoctorShare: shares.doctorShareRaw, DoctorShareNote: '', ClinicShare: shares.clinicShareRaw,
        VisitStatus: record.visitStatus === 'completed' ? 'completed' : 'scheduled',
        UsePromo: record.usePromo ? 'TRUE' : 'FALSE', PromoJSON: JSON.stringify(promoOut), PromoAmount: promoAmount,
        UseMedicine: record.useMedicine ? 'TRUE' : 'FALSE', MedicineJSON: JSON.stringify(medicineGroup.items), MedicineAmount: medicineGroup.total,
        UseOther: record.useOther ? 'TRUE' : 'FALSE', OtherItemsJSON: JSON.stringify(otherGroup.items),
        SupplierPayAmount: supplierPayAmount,
        // Round 14 item 1.7 — OPD photo attachment (Drive fileId, uploaded
        // client-side via the shared resumable-chunk-upload flow before
        // save; see startResumableUpload's 'opd-photo' kind).
        OpdImageFileId: record.opdImageFileId || '',
        UseSupplies: supplies.length ? 'TRUE' : 'FALSE',
        UseOwnMed: ownMedItems.length ? 'TRUE' : 'FALSE', OwnMedJSON: JSON.stringify(ownMedItems),
        UseDepositUse: depositUseItems.length ? 'TRUE' : 'FALSE', DepositUseJSON: JSON.stringify(depositUseItems),
        UseMessenger: messengerItems.length ? 'TRUE' : 'FALSE', MessengerJSON: JSON.stringify(messengerItems),
        StockApplied: isNew ? 'TRUE' : (existing.StockApplied || '')
      };

      // Stock + ยาฝาก first: it throws (before anything is written) if
      // stock or a deposit balance would go negative.
      applyRecordEffects_(session, isNew ? { stock: {}, deposit: {} } : recordEffects_(existing), recordEffects_(obj),
        'บันทึกการรักษา ' + obj.RID + ' · ' + obj.PatientName, obj.PatientName);

      if (isNew) {
        appendRow_(sheet, CONFIG.HEADERS.Records, obj);
      } else {
        // Round 19 item 4: `existing` (fetched once above via `existingRows`)
        // already carries `_row` — every getRows_() row is stamped with its
        // 1-based sheet row number — so reuse it directly instead of a
        // second full-sheet findRowIndexById_() read+linear-scan for the
        // exact same RID we already looked up a few lines earlier. At
        // 5000+ Records this alone removes one full-sheet read per edit.
        updateRowByIndex_(sheet, CONFIG.HEADERS.Records, existing._row, obj);
      }

      // Round 17 item 2: if this edit replaced or removed the OPD photo
      // (the new opdImageFileId differs from what was stored before), trash
      // the OLD Drive file now that the new value is safely written — "ถ้ามี
      // การแก้ไขรูปให้ลบรูปเดิมก่อน แล้วบันทึกรูปใหม่". Done AFTER the sheet
      // write (not before) so a Drive failure here never blocks the save
      // itself; wrapped defensively since the old file may already be gone
      // or inaccessible. Never trashes on a brand-new record (isNew || no
      // `existing`) since there's nothing old to clean up.
      if (!isNew && existing && existing.OpdImageFileId && existing.OpdImageFileId !== obj.OpdImageFileId) {
        try { DriveApp.getFileById(existing.OpdImageFileId).setTrashed(true); } catch (e) { /* already gone/inaccessible — nothing to clean up */ }
      }

      // Round 19 item 4: read the Records sheet ONCE now that the write
      // above is done, and hand that single snapshot to both
      // recomputeDoctorDayShares_ calls below instead of each one doing its
      // own independent full-sheet getRows_() read — appendRow_/
      // updateRowByIndex_ never shift existing row indices (only a delete
      // does), so this snapshot's `_row` numbers stay valid for both calls.
      // The old-group and new-group filters are always disjoint (they only
      // differ when doctorId/date actually changed), so reusing one
      // snapshot for both is equivalent to two independent fresh reads.
      var postWriteRows = getRows_(sheet);
      // Re-derive the doctor's daily-minimum-guarantee top-up across every
      // record for this doctor+date (including the one just written above).
      var newGroup = recomputeDoctorDayShares_(doctor.DoctorID, record.date, postWriteRows);
      // If this was an edit that moved the record to a different doctor
      // and/or date, the OLD group lost a record and must also be
      // recalculated (the shortfall/top-up there may now look different).
      if (!isNew && (oldDoctorId !== doctor.DoctorID || oldDate !== record.date)) {
        recomputeDoctorDayShares_(oldDoctorId, oldDate, postWriteRows);
      }

      writeAudit_(session, isNew ? 'CREATE' : 'UPDATE', 'Record', obj.RID, obj.PatientName + ' ' + obj.Amount);
      bumpDashboardCacheVersion_(); // Round 19 item 4: treatment records feed dashboard totals
      // The record we just saved always lands inside `newGroup` — it's
      // exactly what recomputeDoctorDayShares_ filtered for
      // (doctor.DoctorID + record.date) — so read its final
      // DoctorShare/ClinicShare straight off that in-memory result instead
      // of a second full-sheet re-read+filter to see what the top-up
      // mechanism actually wrote.
      var savedRow = newGroup.filter(function (r) { return r.RID === obj.RID; })[0];
      return { ok: true, data: mapRecordOut_(savedRow || Object.assign({ _row: 0 }, obj)) };
    } finally {
      lock.releaseLock();
    }
  });
}

function deleteRecord(token, id) {
  return safeCall_(function () {
    var session = requireAuth_(token, 'admin');
    var lock = LockService.getScriptLock();
    lock.waitLock(30000);
    try {
      var sheet = getSheet_(CONFIG.SHEETS.RECORDS);
      var rows = getRows_(sheet);
      var existing = rows.filter(function (r) { return r.RID === id; })[0];
      if (!existing) throw ValidationError_('ไม่พบข้อมูลการรักษา');
      // Put back whatever this record took out of Stock / moved in ยาฝาก
      // (throws, deleting nothing, if its deposited medicine was already used).
      applyRecordEffects_(session, recordEffects_(existing), { stock: {}, deposit: {} }, 'ลบบันทึกการรักษา ' + id + ' · ' + existing.PatientName, existing.PatientName);
      var doctorId = existing.DoctorID, date = dateStr_(existing.Date);
      deleteRowByIndex_(sheet, existing._row);
      // The deleted record's share of the daily minimum (if any) needs to
      // be re-distributed among whatever's left in the group.
      recomputeDoctorDayShares_(doctorId, date);
      // Round 17 item 2 (natural extension of the same cleanup): deleting
      // the record entirely should also clean up its attached OPD photo in
      // Drive, same reasoning as the edit-time cleanup in saveRecord —
      // otherwise deleted records leave permanently orphaned files behind.
      if (existing.OpdImageFileId) {
        try { DriveApp.getFileById(existing.OpdImageFileId).setTrashed(true); } catch (e) { /* already gone/inaccessible */ }
      }
      writeAudit_(session, 'DELETE', 'Record', id, '');
      bumpDashboardCacheVersion_(); // Round 19 item 4: treatment records feed dashboard totals
      return { ok: true };
    } finally {
      lock.releaseLock();
    }
  });
}

/**
 * Issues (or re-issues) a printable receipt for a treatment record:
 * assigns a REC-YYYYMMDD-XXX number (once — reuses the existing number on
 * reprint), stores a short-lived one-time view ticket in CacheService, and
 * returns the receipt URL. Receipt.html's visual design was finalized in a
 * later round (the clinic supplied the exact layout/theme to use).
 */
function issueReceipt(token, rid) {
  return safeCall_(function () {
    var session = requireAuth_(token, 'staff');
    var sheet = getSheet_(CONFIG.SHEETS.RECORDS);
    var rows = getRows_(sheet);
    var record = rows.filter(function (r) { return r.RID === rid; })[0];
    if (!record) throw ValidationError_('ไม่พบข้อมูลการรักษา');

    var receiptNo = record.ReceiptNo;
    if (!receiptNo) {
      receiptNo = generateReceiptNo_();
      updateRowByIndex_(sheet, CONFIG.HEADERS.Records, record._row, Object.assign({}, record, { ReceiptNo: receiptNo, UpdatedAt: nowStr_() }));
    }

    var ticket = Utilities.getUuid();
    CacheService.getScriptCache().put(CONFIG.CACHE.RECEIPT_TICKET_PREFIX + ticket, JSON.stringify({ rid: rid }), CONFIG.CACHE.RECEIPT_TICKET_TTL);
    var url = ScriptApp.getService().getUrl() + '?page=receipt&ticket=' + ticket;
    writeAudit_(session, 'ISSUE_RECEIPT', 'Record', rid, receiptNo);
    return { ok: true, data: { receiptNo: receiptNo, url: url, ticket: ticket } };
  });
}

/** Builds the data object handed to Receipt.html for rendering / printing. */
function buildReceiptData_(rid) {
  var rows = getRows_(getSheet_(CONFIG.SHEETS.RECORDS));
  var r = rows.filter(function (x) { return x.RID === rid; })[0];
  if (!r) return null;
  var settingsRows = getRows_(getSheet_(CONFIG.SHEETS.CLINIC_SETTINGS));
  var clinicInfo = {};
  settingsRows.forEach(function (row) { clinicInfo[row.Key] = row.Value; });
  // Inline the logo as base64 for the same reason as getLoginImageUrl_ above
  // (Receipt.html is rendered server-side via doGet, so this is the one
  // spot that can fetch the blob before the page is ever sent to a browser).
  clinicInfo.logoUrl = '';
  if (clinicInfo.logoFileId) {
    try {
      var logoBlob = DriveApp.getFileById(clinicInfo.logoFileId).getBlob();
      clinicInfo.logoUrl = 'data:' + logoBlob.getContentType() + ';base64,' + Utilities.base64Encode(logoBlob.getBytes());
    } catch (e) { /* missing/inaccessible file — receipt still renders without a logo */ }
  }
  var supplies = parseSupplies_(r.SuppliesJSON);
  // Round 17 item 1: matches the canonical check used everywhere else in
  // this file (mapRecordOut_, recordTotalCollected_, computeDailySummaryDetail_,
  // buildRecordShares_'s save path, etc. — search UseCanula) — this was the
  // ONLY UseCanula check in the whole codebase that didn't also accept a raw
  // boolean `true` (only the string 'TRUE'). SpreadsheetApp can hand back a
  // real boolean for a cell depending on how it was written/formatted, so a
  // record with canula actually enabled could silently print with NO canula
  // line on the receipt while still showing correctly everywhere else in the
  // app (record panel, day panel, LINE push, dashboard) — very likely the
  // exact bug behind "เพิ่มยอด canula ด้วย".
  var useCanula = String(r.UseCanula) === 'TRUE' || r.UseCanula === true;
  var canulaAmount = Number(r.CanulaAmount) || 0;
  var otherAmount = Number(r.OtherAmount) || 0;
  var amount = Number(r.Amount) || 0;
  var promoAmount = Number(r.PromoAmount) || 0;
  var medicineAmount = Number(r.MedicineAmount) || 0;
  return {
    receiptNo: r.ReceiptNo, date: dateStr_(r.Date), patientId: r.PatientID || '', patientName: r.PatientName, phone: r.Phone,
    serviceCategory: r.ServiceCategory, detail: r.Detail, doctorName: r.DoctorName, paymentMethod: r.PaymentMethod,
    note: r.Note || '',
    supplies: supplies, useCanula: useCanula, canulaAmount: canulaAmount, otherAmount: otherAmount,
    // Round 13 itemized groups — see Records.PromoJSON/MedicineJSON/
    // OtherItemsJSON schema comment. Printed as their own receipt lines
    // alongside the existing supplies/canula lines.
    promoItems: parseLineItems_(r.PromoJSON), promoAmount: promoAmount,
    medicineItems: parseLineItems_(r.MedicineJSON), medicineAmount: medicineAmount,
    otherItems: parseLineItems_(r.OtherItemsJSON),
    // `amount` is จำนวนเงินหัตถการ (base procedure price, no longer includes
    // canula — see the revenue-split engine); `grandTotal` is what actually
    // gets printed as the amount due, matching totalCollected in mapRecordOut_.
    amount: amount, grandTotal: round2_(amount + promoAmount + (useCanula ? canulaAmount : 0) + medicineAmount + otherAmount),
    clinic: clinicInfo
  };
}

/* ======================================================================
   บันทึกค่าใช้จ่ายคลีนิค — Expenses (tab: ค่าใช้จ่าย) + Orders/OrderLines (tab: จัดซื้อ)
   ====================================================================== */

function mapExpenseOut_(e) {
  return { id: e.ExpID, date: dateStr_(e.Date), name: e.Name, category: e.Category, amount: Number(e.Amount) || 0, createdBy: e.CreatedBy, row: e._row };
}

function listExpenses(token, opts) {
  return safeCall_(function () {
    requireAuth_(token, 'staff');
    opts = opts || {};
    var filters = compactFilters_(opts.filters);
    var rows = getRows_(getSheet_(CONFIG.SHEETS.EXPENSES));
    if (filters.q) {
      var expQ = String(filters.q).trim().toLowerCase();
      rows = rows.filter(function (r) { return String(r.Name).toLowerCase().indexOf(expQ) > -1; });
    }
    if (filters.category) rows = rows.filter(function (r) { return r.Category === filters.category; });
    if (filters.dateFrom) rows = rows.filter(function (r) { return dateStr_(r.Date) >= filters.dateFrom; });
    if (filters.dateTo) rows = rows.filter(function (r) { return dateStr_(r.Date) <= filters.dateTo; });
    // Same date-sort fix as listRecords (round 12) — compare normalized date
    // strings, not raw Date objects.
    rows.sort(function (a, b) { return dateStr_(b.Date).localeCompare(dateStr_(a.Date)); });
    var page = paginate_(rows, opts.page, opts.pageSize);
    page.items = page.items.map(mapExpenseOut_);
    return { ok: true, data: page };
  });
}

function saveExpense(token, expense) {
  return safeCall_(function () {
    var session = requireAuth_(token, 'staff');
    requireFields_(expense, { date: 'วันที่', name: 'รายการ', category: 'หมวดหมู่' });
    requireDateStr_(expense.date, 'วันที่');
    var amount = requireNumber_(expense.amount || 0, 'จำนวนเงิน', { min: 0 });
    var lock = LockService.getScriptLock();
    lock.waitLock(30000);
    try {
      var sheet = getSheet_(CONFIG.SHEETS.EXPENSES);
      var isNew = !expense.id;
      var obj = {
        ExpID: isNew ? ('EX' + nextSeq_('EXPENSE', 5)) : expense.id, Date: expense.date, Name: expense.name,
        Category: expense.category, Amount: amount, CreatedBy: isNew ? session.username : (expense.createdBy || session.username),
        CreatedAt: isNew ? nowStr_() : (expense.createdAt || nowStr_())
      };
      if (isNew) appendRow_(sheet, CONFIG.HEADERS.Expenses, obj);
      else {
        var rowIndex = findRowIndexById_(sheet, 'ExpID', expense.id);
        if (rowIndex < 0) throw ValidationError_('ไม่พบรายการค่าใช้จ่าย');
        updateRowByIndex_(sheet, CONFIG.HEADERS.Expenses, rowIndex, obj);
      }
      writeAudit_(session, isNew ? 'CREATE' : 'UPDATE', 'Expense', obj.ExpID, obj.Name + ' ' + obj.Amount);
      bumpDashboardCacheVersion_(); // Round 19 item 4: expenses feed dashboard totals
      return { ok: true, data: mapExpenseOut_(Object.assign({ _row: 0 }, obj)) };
    } finally {
      lock.releaseLock();
    }
  });
}

function deleteExpense(token, id) {
  return safeCall_(function () {
    var session = requireAuth_(token, 'admin');
    var sheet = getSheet_(CONFIG.SHEETS.EXPENSES);
    var rowIndex = findRowIndexById_(sheet, 'ExpID', id);
    if (rowIndex < 0) throw ValidationError_('ไม่พบรายการค่าใช้จ่าย');
    deleteRowByIndex_(sheet, rowIndex);
    writeAudit_(session, 'DELETE', 'Expense', id, '');
    bumpDashboardCacheVersion_(); // Round 19 item 4: expenses feed dashboard totals
    return { ok: true };
  });
}

/* ---- จัดซื้อ (purchase orders): Orders (header) + OrderLines (multi-line items) ---- */

function mapOrderOut_(o, lines) {
  return {
    id: o.PoID, date: dateStr_(o.Date), supplier: o.Supplier, totalAmount: Number(o.TotalAmount) || 0,
    createdBy: o.CreatedBy, row: o._row,
    lines: (lines || []).map(function (l) {
      return { id: l.LineID, category: l.Category, itemName: l.ItemName, qty: Number(l.Qty) || 0, unitPrice: Number(l.UnitPrice) || 0, lineTotal: Number(l.LineTotal) || 0 };
    })
  };
}

function listOrders(token, opts) {
  return safeCall_(function () {
    requireAuth_(token, 'staff');
    opts = opts || {};
    var filters = compactFilters_(opts.filters);
    var rows = getRows_(getSheet_(CONFIG.SHEETS.ORDERS));
    // A purchase order (PO) is a header row with its actual items on separate
    // OrderLines rows (multi-line, possibly spanning several categories) — so
    // "search by item name" / "filter by category" both have to look at the
    // OrderLines sheet and then map matching line(s) back to their PoID,
    // rather than filtering the Orders sheet's own columns directly.
    var lineRows = (filters.q || filters.category) ? getRows_(getSheet_(CONFIG.SHEETS.ORDER_LINES)) : null;
    if (filters.category) {
      var catPoIds = {};
      lineRows.forEach(function (l) { if (l.Category === filters.category) catPoIds[l.PoID] = true; });
      rows = rows.filter(function (r) { return !!catPoIds[r.PoID]; });
    }
    if (filters.q) {
      var q = String(filters.q).trim().toLowerCase();
      var itemPoIds = {};
      lineRows.forEach(function (l) { if (String(l.ItemName).toLowerCase().indexOf(q) > -1) itemPoIds[l.PoID] = true; });
      // Matches either the supplier name OR any line item's name within that PO.
      rows = rows.filter(function (r) { return String(r.Supplier).toLowerCase().indexOf(q) > -1 || !!itemPoIds[r.PoID]; });
    }
    if (filters.dateFrom) rows = rows.filter(function (r) { return dateStr_(r.Date) >= filters.dateFrom; });
    if (filters.dateTo) rows = rows.filter(function (r) { return dateStr_(r.Date) <= filters.dateTo; });
    // Same date-sort fix as listRecords (round 12) — compare normalized date
    // strings, not raw Date objects.
    rows.sort(function (a, b) { return dateStr_(b.Date).localeCompare(dateStr_(a.Date)); });
    var page = paginate_(rows, opts.page, opts.pageSize);
    page.items = page.items.map(function (o) { return mapOrderOut_(o, []); }); // lines omitted from list view for size; fetched via getOrder
    return { ok: true, data: page };
  });
}

function getOrder(token, id) {
  return safeCall_(function () {
    requireAuth_(token, 'staff');
    var orderRows = getRows_(getSheet_(CONFIG.SHEETS.ORDERS));
    var order = orderRows.filter(function (o) { return o.PoID === id; })[0];
    if (!order) throw ValidationError_('ไม่พบใบสั่งซื้อ');
    var lines = getRows_(getSheet_(CONFIG.SHEETS.ORDER_LINES)).filter(function (l) { return l.PoID === id; });
    return { ok: true, data: mapOrderOut_(order, lines) };
  });
}

/** order: {id?, date, supplier, lines:[{category,itemName,qty,unitPrice}]} — TotalAmount computed server-side. */
function saveOrder(token, order) {
  return safeCall_(function () {
    var session = requireAuth_(token, 'staff');
    requireFields_(order, { date: 'วันที่', supplier: 'ผู้จำหน่าย' });
    requireDateStr_(order.date, 'วันที่');
    var lines = Array.isArray(order.lines) ? order.lines : [];
    if (!lines.length) throw ValidationError_('กรุณาเพิ่มรายการสินค้าอย่างน้อย 1 รายการ');
    lines.forEach(function (l, i) {
      requireFields_(l, { category: 'หมวดหมู่ (บรรทัดที่ ' + (i + 1) + ')', itemName: 'ชื่อสินค้า (บรรทัดที่ ' + (i + 1) + ')' });
      requireNumber_(l.qty, 'จำนวน (บรรทัดที่ ' + (i + 1) + ')', { min: 0.0001 });
      requireNumber_(l.unitPrice, 'ราคาต่อหน่วย (บรรทัดที่ ' + (i + 1) + ')', { min: 0 });
    });
    var totalAmount = lines.reduce(function (sum, l) { return sum + Number(l.qty) * Number(l.unitPrice); }, 0);

    var lock = LockService.getScriptLock();
    lock.waitLock(30000);
    try {
      var orderSheet = getSheet_(CONFIG.SHEETS.ORDERS);
      var lineSheet = getSheet_(CONFIG.SHEETS.ORDER_LINES);
      var isNew = !order.id;
      var poId = isNew ? ('PO' + nextSeq_('ORDER', 5)) : order.id;
      var obj = {
        PoID: poId, Date: order.date, Supplier: order.supplier, TotalAmount: totalAmount,
        CreatedBy: isNew ? session.username : (order.createdBy || session.username),
        CreatedAt: isNew ? nowStr_() : (order.createdAt || nowStr_())
      };
      if (isNew) appendRow_(orderSheet, CONFIG.HEADERS.Orders, obj);
      else {
        var rowIndex = findRowIndexById_(orderSheet, 'PoID', order.id);
        if (rowIndex < 0) throw ValidationError_('ไม่พบใบสั่งซื้อ');
        updateRowByIndex_(orderSheet, CONFIG.HEADERS.Orders, rowIndex, obj);
        // Replace all existing lines for this PO (simplest consistent approach for a small per-PO line count).
        var existingLineRows = getRows_(lineSheet).filter(function (l) { return l.PoID === order.id; });
        existingLineRows.sort(function (a, b) { return b._row - a._row; }); // delete bottom-up to keep row indices valid
        existingLineRows.forEach(function (l) { deleteRowByIndex_(lineSheet, l._row); });
      }
      lines.forEach(function (l, i) {
        appendRow_(lineSheet, CONFIG.HEADERS.OrderLines, {
          LineID: poId + '-' + (i + 1), PoID: poId, Category: l.category, ItemName: l.itemName,
          Qty: Number(l.qty), UnitPrice: Number(l.unitPrice), LineTotal: Number(l.qty) * Number(l.unitPrice)
        });
      });
      writeAudit_(session, isNew ? 'CREATE' : 'UPDATE', 'Order', poId, obj.Supplier + ' ' + totalAmount);
      bumpDashboardCacheVersion_(); // Round 19 item 4: purchase orders feed dashboard expense totals
      return { ok: true, data: mapOrderOut_(Object.assign({ _row: 0 }, obj), lines) };
    } finally {
      lock.releaseLock();
    }
  });
}

function deleteOrder(token, id) {
  return safeCall_(function () {
    var session = requireAuth_(token, 'admin');
    var orderSheet = getSheet_(CONFIG.SHEETS.ORDERS);
    var rowIndex = findRowIndexById_(orderSheet, 'PoID', id);
    if (rowIndex < 0) throw ValidationError_('ไม่พบใบสั่งซื้อ');
    deleteRowByIndex_(orderSheet, rowIndex);
    var lineSheet = getSheet_(CONFIG.SHEETS.ORDER_LINES);
    var lineRows = getRows_(lineSheet).filter(function (l) { return l.PoID === id; });
    lineRows.sort(function (a, b) { return b._row - a._row; });
    lineRows.forEach(function (l) { deleteRowByIndex_(lineSheet, l._row); });
    writeAudit_(session, 'DELETE', 'Order', id, '');
    bumpDashboardCacheVersion_(); // Round 19 item 4: purchase orders feed dashboard expense totals
    return { ok: true };
  });
}

/* ======================================================================
   DASHBOARD — aggregated stats + chart data
   Revenue = sum of Amount on Records with Status = 'paid'.
   Expense = sum of Expenses.Amount + Orders.TotalAmount (all money out).
   NOTE (assumption, flagged for the clinic to confirm): "รายได้เซลล์"
   is revenue attributed to a record that has a SalesID; "ยอดรายได้คลีนิค"
   is revenue with no SalesID (walk-in / clinic-direct); "รายได้อื่นๆ" is
   reserved for a future non-treatment income source and is 0 unless such
   records exist, so the three always sum to total revenue.
   ====================================================================== */

function monthKey_(dateStr) { return String(dateStr).substring(0, 7); } // yyyy-MM

// Round 19 item 4: see the CONFIG.CACHE.DASH_* comment above — a simple
// incrementing counter, read/bumped through CacheService like every other
// cache key in this file (no PropertiesService needed; if the version key
// itself ever expires, dashboardCacheVersion_() just starts back at '0',
// which only means a handful of already-cached entries become unreachable
// early — never a staleness/correctness risk).
function dashboardCacheVersion_() {
  return CacheService.getScriptCache().get(CONFIG.CACHE.DASH_VERSION_KEY) || '0';
}
function bumpDashboardCacheVersion_() {
  var cache = CacheService.getScriptCache();
  var v = (Number(dashboardCacheVersion_()) || 0) + 1;
  cache.put(CONFIG.CACHE.DASH_VERSION_KEY, String(v), CONFIG.CACHE.MASTER_DATA_TTL);
}

function getDashboardData(token, opts) {
  return safeCall_(function () {
    requireAuth_(token, 'staff');
    opts = opts || {};
    var dateFrom = opts.dateFrom || (todayStr_().substring(0, 8) + '01'); // default: this month
    var dateTo = opts.dateTo || todayStr_();

    // Round 19 item 4: getDashboardData used to do 3 full-sheet reads
    // (Records/Expenses/Orders) plus ~8 aggregation passes on EVERY call,
    // including repeat loads of the exact same date range (e.g. switching
    // pages and back, or two staff viewing the dashboard around the same
    // time) — the most expensive unpaginated read in the app, and the one
    // most likely to be felt as "slow" once Records passes 5000+ rows.
    // Cached per date range; see dashboardCacheVersion_ above for why this
    // is safe to serve for up to DASH_TTL without going stale after a
    // relevant write.
    var cache = CacheService.getScriptCache();
    var cacheKey = CONFIG.CACHE.DASH_KEY_PREFIX + dashboardCacheVersion_() + '_' + dateFrom + '_' + dateTo;
    var cachedData = cache.get(cacheKey);
    if (cachedData) return { ok: true, data: JSON.parse(cachedData) };

    var allRecords = getRows_(getSheet_(CONFIG.SHEETS.RECORDS));
    var paidRecords = allRecords.filter(function (r) { return r.Status === 'paid'; });
    var periodRecords = paidRecords.filter(function (r) { var d = dateStr_(r.Date); return d >= dateFrom && d <= dateTo; });

    var allExpenses = getRows_(getSheet_(CONFIG.SHEETS.EXPENSES));
    var periodExpenses = allExpenses.filter(function (r) { var d = dateStr_(r.Date); return d >= dateFrom && d <= dateTo; });
    var allOrders = getRows_(getSheet_(CONFIG.SHEETS.ORDERS));
    var periodOrders = allOrders.filter(function (r) { var d = dateStr_(r.Date); return d >= dateFrom && d <= dateTo; });

    // True total money collected per record = จำนวนเงินหัตถการ + จำนวนเงินโปร +
    // canula (if used) + จำนวนเงินยา + จำนวนเงินอื่นๆ (round 13). This ties out
    // exactly with clinicShare+salesShare+doctorShare+medicineAmount+otherAmount
    // (see recomputeDoctorDayShares_ / computeRecordShares_ — Medicine/Other
    // are deliberately excluded from every revenue share so SupplierPayAmount
    // always balances as its own bucket; promo money IS inside the 3-way split).
    function recordTotalCollected_(r) {
      var canula = (String(r.UseCanula) === 'TRUE' || r.UseCanula === true) ? (Number(r.CanulaAmount) || 0) : 0;
      return (Number(r.Amount) || 0) + (Number(r.PromoAmount) || 0) + canula + (Number(r.MedicineAmount) || 0) + (Number(r.OtherAmount) || 0);
    }

    var totalRevenue = periodRecords.reduce(function (s, r) { return s + recordTotalCollected_(r); }, 0);
    var expenseFromExpenses = periodExpenses.reduce(function (s, r) { return s + (Number(r.Amount) || 0); }, 0);
    var expenseFromOrders = periodOrders.reduce(function (s, r) { return s + (Number(r.TotalAmount) || 0); }, 0);
    var totalExpense = expenseFromExpenses + expenseFromOrders;
    var profit = totalRevenue - totalExpense;
    var profitPct = totalRevenue > 0 ? (profit / totalRevenue) * 100 : 0;

    // Clinic/sales/other revenue now come straight from each record's
    // computed share fields (see the revenue split engine), instead of the
    // earlier placeholder that naively attributed a whole record's amount
    // to "sales revenue" whenever a salesperson was assigned.
    var salesRevenue = periodRecords.reduce(function (s, r) { return s + (Number(r.SalesShare) || 0); }, 0);
    var clinicRevenue = periodRecords.reduce(function (s, r) { return s + (Number(r.ClinicShare) || 0); }, 0);
    // Round 16 item 4: still computed (some downstream reporting may still
    // want it) even though it's no longer one of the 8 dashboard stat cards
    // — see stats.doctorRevenue below, which replaced it in the card list.
    var otherRevenue = periodRecords.reduce(function (s, r) { return s + (Number(r.OtherAmount) || 0); }, 0);
    var totalDoctorShare = periodRecords.reduce(function (s, r) { return s + (Number(r.DoctorShare) || 0); }, 0);

    var uniquePatientIds = {};
    periodRecords.forEach(function (r) { uniquePatientIds[r.PatientID] = true; });
    var customerCount = Object.keys(uniquePatientIds).length;

    // --- Monthly trend (last 12 calendar months ending this month), independent of the period filter ---
    var months = [];
    var now = new Date();
    for (var i = 11; i >= 0; i--) {
      var d = new Date(now.getFullYear(), now.getMonth() - i, 1);
      months.push(Utilities.formatDate(d, Session.getScriptTimeZone() || 'Asia/Bangkok', 'yyyy-MM'));
    }
    var monthlyRevenue = {}, monthlyExpense = {};
    months.forEach(function (m) { monthlyRevenue[m] = 0; monthlyExpense[m] = 0; });
    paidRecords.forEach(function (r) { var m = monthKey_(dateStr_(r.Date)); if (monthlyRevenue[m] !== undefined) monthlyRevenue[m] += recordTotalCollected_(r); });
    allExpenses.forEach(function (r) { var m = monthKey_(dateStr_(r.Date)); if (monthlyExpense[m] !== undefined) monthlyExpense[m] += Number(r.Amount) || 0; });
    allOrders.forEach(function (r) { var m = monthKey_(dateStr_(r.Date)); if (monthlyExpense[m] !== undefined) monthlyExpense[m] += Number(r.TotalAmount) || 0; });
    var monthlyTrend = months.map(function (m) {
      var rev = monthlyRevenue[m], exp = monthlyExpense[m];
      // Round 15 item 8.2: "profit" (บาท) is new, additive — the trend chart
      // now plots the profit LINE against the same money axis as the
      // revenue/expense BARS instead of a separate %-based right axis, so it
      // needs the raw money amount, not just profitPct.
      return { month: m, revenue: rev, expense: exp, profit: rev - exp, profitPct: rev > 0 ? ((rev - exp) / rev) * 100 : 0 };
    });

    // --- Revenue share by type: clinic / doctor / sales split of collected
    // record revenue, plus total supplier spend (Orders) as its own slice —
    // replaces the old by-treatment-category breakdown (round 12, item 1). ---
    var byRevenueType = {
      'ส่วนแบ่งคลีนิค': clinicRevenue,
      'ส่วนแบ่งแพทย์': totalDoctorShare,
      'ส่วนแบ่งเซลล์': salesRevenue,
      'ยอดจ่าย Supplier': expenseFromOrders
    };

    // --- Revenue by doctor (their actual commission share, not the full record amount) ---
    var byDoctor = {};
    periodRecords.forEach(function (r) { byDoctor[r.DoctorName] = (byDoctor[r.DoctorName] || 0) + (Number(r.DoctorShare) || 0); });

    // --- Revenue by sales (their actual commission share) ---
    var bySales = {};
    periodRecords.filter(function (r) { return r.SalesID; }).forEach(function (r) { bySales[r.SalesName] = (bySales[r.SalesName] || 0) + (Number(r.SalesShare) || 0); });

    // --- Expense by category (Expenses sheet categories + a single lump "จัดซื้อ" bucket for Orders) ---
    var byExpenseCat = {};
    periodExpenses.forEach(function (r) { byExpenseCat[r.Category] = (byExpenseCat[r.Category] || 0) + (Number(r.Amount) || 0); });
    if (expenseFromOrders > 0) byExpenseCat['จัดซื้อ'] = (byExpenseCat['จัดซื้อ'] || 0) + expenseFromOrders;

    // --- Revenue by promotion (Round 16 item 5) — each รายการโปร line's own
    // `amount`, summed by promotion name across the period's paid records.
    // PromoJSON is only ever non-empty when a record actually used a promo
    // (see saveRecord's usePromo gate), so no extra UsePromo check is needed
    // here — same pattern as the existing byDoctor/bySales aggregations. ---
    var byPromotion = {};
    periodRecords.forEach(function (r) {
      parseLineItems_(r.PromoJSON).forEach(function (p) {
        var name = p.name || 'ไม่ระบุ';
        byPromotion[name] = (byPromotion[name] || 0) + (Number(p.amount) || 0);
      });
    });

    function toChartArray(map) {
      return Object.keys(map).map(function (k) { return { label: k, value: map[k] }; }).sort(function (a, b) { return b.value - a.value; });
    }

    var result = {
      period: { from: dateFrom, to: dateTo },
      stats: {
        totalRevenue: totalRevenue, totalExpense: totalExpense, profit: profit, profitPct: profitPct,
        customerCount: customerCount, clinicRevenue: clinicRevenue, salesRevenue: salesRevenue, otherRevenue: otherRevenue,
        // Round 16 item 4: stats.doctorRevenue (a money total) now feeds
        // the 8th dashboard stat card ("ยอดรายได้ส่วนแบ่งแพทย์"), replacing
        // otherRevenue there — same "same key name, different shape, in
        // stats-vs-top-level" pattern already used by clinicRevenue/
        // salesRevenue (a money total in `stats`, a chart array below).
        doctorRevenue: totalDoctorShare
      },
      monthlyTrend: monthlyTrend,
      revenueByType: toChartArray(byRevenueType),
      doctorRevenue: toChartArray(byDoctor),
      salesRevenue: toChartArray(bySales),
      expenseByCategory: toChartArray(byExpenseCat),
      promoRevenue: toChartArray(byPromotion)
    };
    try { cache.put(cacheKey, JSON.stringify(result), CONFIG.CACHE.DASH_TTL); } catch (e) { /* too large for cache (very wide monthlyTrend/chart data) — still return live data */ }
    return { ok: true, data: result };
  });
}

/* ======================================================================
   DRIVE — resumable chunk upload (large files, e.g. patient photos)
   Client flow:
     1. startResumableUpload(token, opts) -> {sessionUrl, fileName}
     2. uploadFileChunk(token, sessionUrl, base64Chunk, rangeStart, rangeEnd, totalSize)
        repeatedly for each chunk (e.g. 4MB slices) until {done:true, fileId}.
   The OAuth token never reaches the client — every PUT is proxied through
   this server function, which is what makes it safe to expose publicly.
   ====================================================================== */

function guessExtension_(mimeType) {
  var map = { 'image/jpeg': '.jpg', 'image/jpg': '.jpg', 'image/png': '.png', 'image/webp': '.webp', 'application/pdf': '.pdf' };
  return map[mimeType] || '';
}

/** opts: {kind:'photo'|'login-image'|'other', patientId, patientName, fileName, mimeType, fileSize} */
function startResumableUpload(token, opts) {
  return safeCall_(function () {
    var ADMIN_ONLY_KINDS = { 'login-image': 1, 'logo': 1 };
    var session = requireAuth_(token, opts && ADMIN_ONLY_KINDS[opts.kind] ? 'admin' : 'staff');
    opts = opts || {};
    if (!opts.mimeType) throw ValidationError_('ไม่พบชนิดไฟล์');
    var folder, fileName;
    if (opts.kind === 'photo') {
      requireFields_(opts, { patientId: 'รหัสคนไข้', patientName: 'ชื่อคนไข้' });
      folder = getPhotosFolder_();
      // Round 17 item 3: "Customer_<ID>_<Name>" (was "<ID>-<Name>").
      fileName = 'Customer_' + opts.patientId + '_' + opts.patientName + guessExtension_(opts.mimeType);
    } else if (opts.kind === 'login-image') {
      folder = getPhotosFolder_();
      fileName = 'login-image-' + new Date().getTime() + guessExtension_(opts.mimeType);
    } else if (opts.kind === 'logo') {
      folder = getPhotosFolder_();
      fileName = 'logo-' + new Date().getTime() + guessExtension_(opts.mimeType);
    } else if (opts.kind === 'opd-photo') {
      // Round 14 item 1.7 — treatment-record OPD photo. recordId may be
      // 'TEMP' if the record hasn't been saved yet (same pattern as a new
      // patient's photo upload before the patient has an ID).
      // Round 17 item 2: filename is now "opd_<yyyyMMdd>_<recordId>" (was
      // "opd-<recordId>-<timestamp>"), and the file is saved into a
      // Year/Month/Day folder chain (getOpdPhotoDayFolder_) instead of the
      // flat Photos folder — both keyed off the treatment record's own date
      // (opts.date, the modal's "วันที่" field), not today's date, so a
      // backdated record files under the day it's actually for. Falls back
      // to today if the date is missing/unparseable (e.g. a very old client
      // build that doesn't send it yet).
      var opdDate = opts.date ? new Date(opts.date) : new Date();
      if (isNaN(opdDate.getTime())) opdDate = new Date();
      var opdYmd = Utilities.formatDate(opdDate, Session.getScriptTimeZone() || 'Asia/Bangkok', 'yyyyMMdd');
      folder = getOpdPhotoDayFolder_(opdDate);
      fileName = 'opd_' + opdYmd + '_' + (opts.recordId || 'TEMP') + guessExtension_(opts.mimeType);
    } else {
      folder = getPhotosFolder_();
      fileName = opts.fileName || ('upload-' + Utilities.getUuid());
    }
    var url = 'https://www.googleapis.com/upload/drive/v3/files?uploadType=resumable&supportsAllDrives=true';
    var metadata = { name: fileName, parents: [folder.getId()] };
    var resp = UrlFetchApp.fetch(url, {
      method: 'post',
      contentType: 'application/json; charset=UTF-8',
      headers: {
        Authorization: 'Bearer ' + ScriptApp.getOAuthToken(),
        'X-Upload-Content-Type': opts.mimeType,
        'X-Upload-Content-Length': String(opts.fileSize || '')
      },
      payload: JSON.stringify(metadata),
      muteHttpExceptions: true
    });
    if (resp.getResponseCode() !== 200) throw new Error('เริ่มต้นอัปโหลดไม่สำเร็จ: ' + resp.getContentText());
    var headers = resp.getHeaders();
    var sessionUrl = headers['Location'] || headers['location'];
    writeAudit_(session, 'UPLOAD_START', 'Drive', fileName, opts.kind || '');
    return { ok: true, data: { sessionUrl: sessionUrl, fileName: fileName } };
  });
}

/** base64Chunk: one slice of the file, base64-encoded. rangeStart/rangeEnd are 0-based byte offsets (inclusive). */
function uploadFileChunk(token, sessionUrl, base64Chunk, rangeStart, rangeEnd, totalSize) {
  return safeCall_(function () {
    requireAuth_(token, 'staff');
    if (!sessionUrl) throw ValidationError_('ไม่พบ session การอัปโหลด');
    var bytes = Utilities.base64Decode(base64Chunk);
    var resp = UrlFetchApp.fetch(sessionUrl, {
      method: 'put',
      headers: { 'Content-Range': 'bytes ' + rangeStart + '-' + rangeEnd + '/' + totalSize },
      payload: bytes,
      muteHttpExceptions: true
    });
    var code = resp.getResponseCode();
    if (code === 200 || code === 201) {
      var fileInfo = JSON.parse(resp.getContentText());
      return { ok: true, data: { done: true, fileId: fileInfo.id, fileName: fileInfo.name } };
    }
    if (code === 308) return { ok: true, data: { done: false } }; // "Resume Incomplete" — chunk accepted, keep going
    throw new Error('อัปโหลดไม่สำเร็จ (code ' + code + '): ' + resp.getContentText());
  });
}

function ensureFileViewable_(fileId) {
  var file = DriveApp.getFileById(fileId);
  try { file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW); } catch (e) { /* domain policy may forbid; ignore */ }
  return file;
}

/**
 * Returns a photo as an inline base64 data: URI instead of a public Drive
 * link. This is deliberate: `uc?export=view` links only render for a
 * viewer if the file is actually shared "Anyone with the link" — and
 * ensureFileViewable_()'s setSharing() call silently no-ops under a
 * Google Workspace domain policy that restricts external sharing (common
 * in real clinics), which was making uploaded photos fail to display for
 * any staff member who isn't the file's owner even though the upload
 * itself succeeded. Fetching the blob server-side and inlining it sidesteps
 * Drive sharing entirely — it works purely off the deploying account's own
 * Drive access (Execute as: Me), regardless of link-sharing settings.
 * Trade-off: no browser caching across requests and a ~33% base64 size
 * overhead, acceptable for clinic-sized patient photos.
 */
function getPhotoUrl(token, fileId) {
  return safeCall_(function () {
    requireAuth_(token, 'staff');
    if (!fileId) return { ok: true, data: '' };
    var blob = DriveApp.getFileById(fileId).getBlob();
    return { ok: true, data: 'data:' + blob.getContentType() + ';base64,' + Utilities.base64Encode(blob.getBytes()) };
  });
}

/**
 * Server-side (unauthenticated) lookup used only by doGet() to inject the
 * login screen's custom image — the visitor has no session token yet at
 * that point, so this must NOT go through requireAuth_. Always fails soft
 * (returns '') so a bad/missing image or an un-setup project never breaks
 * the login page itself.
 */
function getLoginImageUrl_() {
  try {
    var rows = getRows_(getSheet_(CONFIG.SHEETS.CLINIC_SETTINGS));
    var row = rows.filter(function (r) { return r.Key === 'loginImageFileId'; })[0];
    if (!row || !row.Value) return '';
    // Base64-inline for the same reason as getPhotoUrl() above — avoids
    // depending on Drive link-sharing, which a Workspace domain policy can
    // silently block. Also cheaper here: doGet() already has the blob in
    // hand, so this avoids a second round-trip the client would otherwise
    // need to make just to fetch the login background image.
    var blob = DriveApp.getFileById(row.Value).getBlob();
    return 'data:' + blob.getContentType() + ';base64,' + Utilities.base64Encode(blob.getBytes());
  } catch (e) {
    return '';
  }
}

/**
 * Shared admin-only "replace an image stored as a ClinicSettings key/value
 * pair" logic — used for both the login-screen image and the clinic logo.
 * Trashes the previous Drive file (if any) so replacing an image doesn't
 * leave orphaned files behind, and returns the new file as an inline
 * base64 data: URI (see getPhotoUrl's comment above for why — this used to
 * return a public `uc?export=view` link here specifically, which is the
 * same fragile-under-Workspace-sharing-policy pattern that was fixed
 * everywhere else in Round 4; fixed here too for consistency).
 */
function setClinicImage_(token, key, fileId) {
  var session = requireAuth_(token, 'admin');
  if (!fileId) throw ValidationError_('ไม่พบไฟล์ที่อัปโหลด');
  var sheet = getSheet_(CONFIG.SHEETS.CLINIC_SETTINGS);
  var rows = getRows_(sheet);
  var existing = rows.filter(function (r) { return r.Key === key; })[0];
  var oldFileId = existing ? existing.Value : '';
  if (existing) updateRowByIndex_(sheet, CONFIG.HEADERS.ClinicSettings, existing._row, { Key: key, Value: fileId });
  else appendRow_(sheet, CONFIG.HEADERS.ClinicSettings, { Key: key, Value: fileId });
  if (oldFileId && oldFileId !== fileId) {
    try { DriveApp.getFileById(oldFileId).setTrashed(true); } catch (e) { /* ignore — old file may already be gone */ }
  }
  var blob = DriveApp.getFileById(fileId).getBlob();
  writeAudit_(session, 'UPDATE', 'ClinicSettings', key, fileId);
  invalidateMasterCache_();
  return 'data:' + blob.getContentType() + ';base64,' + Utilities.base64Encode(blob.getBytes());
}
/** Admin-only: removes an image stored as a ClinicSettings key/value pair. */
function removeClinicImage_(token, key) {
  var session = requireAuth_(token, 'admin');
  var sheet = getSheet_(CONFIG.SHEETS.CLINIC_SETTINGS);
  var rows = getRows_(sheet);
  var existing = rows.filter(function (r) { return r.Key === key; })[0];
  if (existing) {
    try { DriveApp.getFileById(existing.Value).setTrashed(true); } catch (e) { /* ignore */ }
    deleteRowByIndex_(sheet, existing._row);
  }
  writeAudit_(session, 'DELETE', 'ClinicSettings', key, '');
  invalidateMasterCache_();
}

/** Admin-only: sets (or replaces) the login screen's custom image, trashing the previous Drive file if any. */
function setLoginImage(token, fileId) {
  return safeCall_(function () {
    var url = setClinicImage_(token, 'loginImageFileId', fileId);
    return { ok: true, data: { url: url } };
  });
}

/** Admin-only: removes the login screen's custom image (falls back to the default look). */
function removeLoginImage(token) {
  return safeCall_(function () {
    removeClinicImage_(token, 'loginImageFileId');
    return { ok: true };
  });
}

/** Admin-only: sets (or replaces) the clinic logo — shown in the sidebar and on printed receipts. */
function setLogo(token, fileId) {
  return safeCall_(function () {
    var url = setClinicImage_(token, 'logoFileId', fileId);
    return { ok: true, data: { url: url } };
  });
}

/** Admin-only: removes the clinic logo. */
function removeLogo(token) {
  return safeCall_(function () {
    removeClinicImage_(token, 'logoFileId');
    return { ok: true };
  });
}

/* ======================================================================
   DEMO DATA — fills every sheet with realistic sample data so the app
   can be explored end-to-end before real clinic data is entered.
   Safe to re-run: clears existing rows (except headers) first.
   Callable from the client (admin only) or run manually in the editor
   (call with no arguments).
   ====================================================================== */

function clearSheetData_(sheet) {
  var lastRow = sheet.getLastRow();
  if (lastRow > 1) sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn()).clearContent();
}

function seedDemoData(token) {
  return safeCall_(function () {
    if (token) requireAuth_(token, 'admin');
    var ss = getSS_();

    Object.keys(CONFIG.SHEETS).forEach(function (key) {
      var name = CONFIG.SHEETS[key];
      if (name === CONFIG.SHEETS.STAFF) return; // keep real login accounts intact
      clearSheetData_(getSheet_(name));
    });

    // Categories
    var catSheet = getSheet_(CONFIG.SHEETS.CATEGORIES);
    var treatmentCatNames = ['โบท็อกซ์', 'ฟิลเลอร์', 'เลเซอร์', 'ดูแลผิวหน้า', 'ร้อยไหม'];
    var stockCatNames = ['ยาและเวชภัณฑ์', 'อุปกรณ์ทางการแพทย์', 'เครื่องสำอาง', 'วัสดุสิ้นเปลือง'];
    var expenseCatNames = ['ค่าเช่า', 'ค่าน้ำค่าไฟ', 'เงินเดือนพนักงาน', 'การตลาด', 'ค่าซ่อมบำรุง', 'อื่นๆ'];
    var catIds = { treatment: [], stock: [], expense: [] };
    [['treatment', treatmentCatNames], ['stock', stockCatNames], ['expense', expenseCatNames]].forEach(function (pair) {
      pair[1].forEach(function (name, i) {
        var id = 'C' + nextSeq_('CATEGORY', 3);
        appendRow_(catSheet, CONFIG.HEADERS.Categories, { CategoryID: id, Type: pair[0], Name: name, SortOrder: i + 1 });
        catIds[pair[0]].push({ id: id, name: name });
      });
    });

    // Doctors — MinAmount is a real (non-zero) daily minimum guarantee here
    // so the seeded data actually exercises the top-up logic in
    // recomputeDoctorDayShares_.
    var docSheet = getSheet_(CONFIG.SHEETS.DOCTORS);
    var doctorSeed = [
      { name: 'พญ.วลิน ศรีสุข', minAmount: 1500, commissionPct: 15 },
      { name: 'นพ.ธนกร รุ่งเรือง', minAmount: 1200, commissionPct: 12 },
      { name: 'พญ.ปิยะดา แสงทอง', minAmount: 1000, commissionPct: 10 }
    ];
    var doctorIds = doctorSeed.map(function (d) {
      var id = 'D' + nextSeq_('DOCTOR', 3);
      appendRow_(docSheet, CONFIG.HEADERS.Doctors, { DoctorID: id, Name: d.name, Status: 'active', MinAmount: d.minAmount, CreatedAt: nowStr_(), CommissionPct: d.commissionPct });
      return { id: id, name: d.name, minAmount: d.minAmount, commissionPct: d.commissionPct };
    });

    // Sales
    var salesSheet = getSheet_(CONFIG.SHEETS.SALES);
    var salesSeed = [
      { name: 'คุณอลิสา มั่งมี', commissionRate: 10 },
      { name: 'คุณกิตติ ขายเก่ง', commissionRate: 8 }
    ];
    var salesIds = salesSeed.map(function (s) {
      var id = 'SL' + nextSeq_('SALES', 3);
      appendRow_(salesSheet, CONFIG.HEADERS.Sales, { SalesID: id, Name: s.name, Status: 'active', CommissionRate: s.commissionRate, MinAmount: 0, CreatedAt: nowStr_() });
      return { id: id, name: s.name, commissionRate: s.commissionRate };
    });

    // Promotions (round 12/13 — added here so seeded data actually exercises
    // the current format: without at least one active promotion, the record
    // modal's "รายการโปร" picker, the dashboard's รายได้รายโปรโมชัน/สัดส่วน
    // charts, and the day panel's "ยอดรายการโปรหน้าคลีนิค" row would all
    // demo as permanently empty. ID prefix ('PM' + nextSeq_('PROMO', 3))
    // matches savePromotion()'s own convention exactly, so this shares the
    // same counter real usage would continue from. One entry is seeded
    // `inactive` so the Promotions page/dropdown-filtering demo both an
    // active and an inactive promotion, same as `r16_check.js` expects.
    var promoSheet = getSheet_(CONFIG.SHEETS.PROMOTIONS);
    var promoSeed = [
      { name: 'โบท็อกซ์หน้าเรียว 50 ยูนิต', price: 8900, doctorSharePct: 20, status: 'active' },
      { name: 'ฟิลเลอร์ใต้ตา 1 CC', price: 12500, doctorSharePct: 25, status: 'active' },
      { name: 'เลเซอร์กระชับผิวหน้า', price: 4500, doctorSharePct: 15, status: 'active' },
      { name: 'โปรแกรมโปรโมชันเก่า', price: 3000, doctorSharePct: 10, status: 'inactive' }
    ];
    var promoIds = promoSeed.map(function (p) {
      var id = 'PM' + nextSeq_('PROMO', 3);
      appendRow_(promoSheet, CONFIG.HEADERS.Promotions, { PromoID: id, Name: p.name, Price: p.price, DoctorSharePct: p.doctorSharePct, Status: p.status, CreatedAt: nowStr_(), UpdatedAt: nowStr_() });
      return { id: id, name: p.name, price: p.price, doctorSharePct: p.doctorSharePct, status: p.status };
    });
    var activePromoIds = promoIds.filter(function (p) { return p.status === 'active'; });

    // Stock
    var stockSheet = getSheet_(CONFIG.SHEETS.STOCK);
    var stockItems = [
      { name: 'โบท็อกซ์ Dysport 500u', category: catIds.stock[0].name, unit: 'ขวด', qty: 25 },
      { name: 'ฟิลเลอร์ Juvederm 1ml', category: catIds.stock[0].name, unit: 'หลอด', qty: 18 },
      { name: 'เข็ม Canula 25G', category: catIds.stock[1].name, unit: 'ชิ้น', qty: 120 },
      { name: 'ถุงมือทางการแพทย์', category: catIds.stock[3].name, unit: 'กล่อง', qty: 40 },
      { name: 'สำลีแอลกอฮอล์', category: catIds.stock[3].name, unit: 'ห่อ', qty: 60 },
      { name: 'ครีมกันแดด SPF50', category: catIds.stock[2].name, unit: 'หลอด', qty: 30 },
      { name: 'เซรั่มวิตามินซี', category: catIds.stock[2].name, unit: 'ขวด', qty: 22 },
      { name: 'ไหมร้อยหน้า PDO', category: catIds.stock[1].name, unit: 'เส้น', qty: 50 }
    ];
    var stockIds = stockItems.map(function (s) {
      var id = 'ST' + nextSeq_('STOCK', 4);
      appendRow_(stockSheet, CONFIG.HEADERS.Stock, { ItemID: id, Name: s.name, Category: s.category, Unit: s.unit, Qty: s.qty, CreatedAt: nowStr_(), UpdatedAt: nowStr_() });
      return Object.assign({ id: id }, s);
    });

    // Patients
    var patSheet = getSheet_(CONFIG.SHEETS.PATIENTS);
    var patientSeed = [
      ['สมหญิง', 'ใจดี', 'มิ้นท์', '0812345678', 'female'], ['อนันต์', 'พูลสวัสดิ์', 'บอส', '0898765432', 'male'],
      ['วิภาวดี', 'รักสวย', 'วิว', '0865551234', 'female'], ['ธีรพงษ์', 'มั่นคง', 'เต้', '0911112222', 'male'],
      ['ชนิดา', 'แสงจันทร์', 'จันทร์', '0823334444', 'female'], ['ปรีชา', 'ทองดี', 'ชา', '0955556666', 'male'],
      ['กมลชนก', 'บุญมี', 'ก้อย', '0877778888', 'female'], ['ณัฐวุฒิ', 'ศรีวิไล', 'นัท', '0899990000', 'male'],
      ['สุภาพร', 'เพชรรัตน์', 'พร', '0801231234', 'female'], ['วรากร', 'ชัยชนะ', 'กร', '0845674567', 'male']
    ];
    var patientIds = patientSeed.map(function (p, i) {
      var id = 'P' + nextSeq_('PATIENT', 5);
      var joined = Utilities.formatDate(new Date(2025, i % 12, (i % 27) + 1), 'Asia/Bangkok', 'yyyy-MM-dd');
      appendRow_(patSheet, CONFIG.HEADERS.Patients, {
        ID: id, FirstName: p[0], LastName: p[1], Phone: p[3], DOB: '', Gender: p[4], Email: '',
        PhotoFileId: '', JoinedDate: joined, CreatedAt: nowStr_(), UpdatedAt: nowStr_(), Nickname: p[2]
      });
      return { id: id, name: p[0] + ' ' + p[1], phone: p[3] };
    });

    // Records — spread across 16 distinct days (each 5 days apart) rather than
    // one record per day: perDayCounts below is deliberately uneven (1 to 4
    // records on a given day) so the calendar/day-summary view always has some
    // days with multiple records to look at, per the round-8 request. Mix of
    // paid/unpaid, with supplies + canula + (round 20) promo + medicine lines
    // on some. Revenue-split fields are computed with the same
    // computeRecordShares_ used by saveRecord, and recomputeDoctorDayShares_
    // is run per doctor+date group afterwards so the seeded data correctly
    // showcases the daily-minimum top-up.
    var recSheet = getSheet_(CONFIG.SHEETS.RECORDS);
    var paymentMethods = ['เงินสด', 'บัตรเครดิต', 'โอนเงิน'];
    var now = new Date();
    var seededGroups = {}; // "doctorId|date" -> {doctorId, date}
    var perDayCounts = [1, 3, 2, 4, 1, 2, 3, 1, 4, 2, 1, 3, 2, 4, 1, 6]; // sums to 40
    var i = 0;
    perDayCounts.forEach(function (count, dayIdx) {
      var recDate = new Date(now.getFullYear(), now.getMonth(), now.getDate());
      recDate.setDate(recDate.getDate() - (dayIdx * 5));
      for (var j = 0; j < count; j++) {
        var patient = patientIds[i % patientIds.length];
        var doctor = doctorIds[i % doctorIds.length];
        var useSales = i % 3 !== 0;
        var sales = useSales ? salesIds[i % salesIds.length] : null;
        var treatCat = catIds.treatment[i % catIds.treatment.length];
        var useCanula = i % 4 === 0;
        var supplies = [];
        if (i % 2 === 0) {
          var stockPick = stockIds[i % stockIds.length];
          supplies.push({ itemId: stockPick.id, name: stockPick.name, qty: 1, unit: stockPick.unit });
        }
        var amount = 1500 + (i % 10) * 850; // ยอดชำระจากลูกค้า (does NOT include canula)
        var canulaAmount = useCanula ? 500 : 0;
        var otherAmount = i % 9 === 0 ? 300 + (i % 5) * 100 : 0; // จำนวนเงินอื่นๆ — separate bucket, not part of anyone's share
        var status = i % 6 === 0 ? 'unpaid' : (i % 17 === 0 ? 'cancelled' : 'paid');
        // Alternate นัดหมาย/เสร็จสิ้น WITHIN a day's group (not just globally by
        // some i%N) so a multi-record day always visibly mixes both statuses.
        var visitStatus = j % 2 === 0 ? 'completed' : 'scheduled';
        // รายการโปร (round 13) — about 1 in 5 records uses one of the active
        // seeded promotions, computed exactly the way saveRecord() derives
        // promoDoctorShare from a promo line, so the demo data actually
        // exercises the promo-aware share formula and gives the dashboard's
        // รายได้รายโปรโมชัน chart / the day panel's item 3 something to show
        // instead of a permanently-empty state.
        var usePromo = activePromoIds.length > 0 && i % 5 === 0;
        var promoLine = null, promoAmount = 0, promoDoctorShare = 0;
        if (usePromo) {
          var promo = activePromoIds[i % activePromoIds.length];
          promoAmount = promo.price;
          promoDoctorShare = round2_(promoAmount * promo.doctorSharePct / 100);
          promoLine = { promoId: promo.id, name: promo.name, price: promo.price, doctorSharePct: promo.doctorSharePct, amount: promoAmount };
        }
        // เบิกยาหน้าคลีนิค (round 13) — about 1 in 7 records, same "not part
        // of anyone's commission share, only feeds SupplierPayAmount" rule
        // as otherAmount above, so day panel item 4/รายการอื่นๆ shows both
        // kinds of line item, not just otherItems.
        var useMedicine = i % 7 === 0;
        var medicineAmount = useMedicine ? 200 : 0;
        var medicineItemsSeed = useMedicine ? [{ name: 'ยาชาเฉพาะที่', amount: medicineAmount }] : [];
        var shares = computeRecordShares_(amount, promoAmount, promoDoctorShare, canulaAmount, { CommissionPct: doctor.commissionPct }, sales ? { CommissionRate: sales.commissionRate } : null);
        appendRow_(recSheet, CONFIG.HEADERS.Records, {
          RID: 'R' + nextSeq_('RECORD', 6), Date: dateStr_(recDate), PatientID: patient.id, PatientName: patient.name,
          Phone: patient.phone, ServiceCategory: treatCat.name, Detail: treatCat.name + ' - โปรแกรมมาตรฐาน',
          DoctorID: doctor.id, DoctorName: doctor.name, SalesID: sales ? sales.id : '', SalesName: sales ? sales.name : '',
          PaymentMethod: paymentMethods[i % paymentMethods.length], SuppliesJSON: JSON.stringify(supplies),
          UseCanula: useCanula ? 'TRUE' : 'FALSE', CanulaAmount: canulaAmount, Amount: amount,
          Status: status, ReceiptNo: '', Note: '', CreatedBy: 'admin', CreatedAt: nowStr_(), UpdatedAt: nowStr_(),
          OtherAmount: otherAmount, SalesCommissionPct: shares.salesPct, SalesShare: shares.salesShare,
          DoctorCommissionPct: shares.doctorPct, DoctorShareRaw: shares.doctorShareRaw, DoctorShare: shares.doctorShareRaw,
          DoctorShareNote: '', ClinicShare: shares.clinicShareRaw, VisitStatus: visitStatus,
          UsePromo: usePromo ? 'TRUE' : 'FALSE', PromoJSON: JSON.stringify(promoLine ? [promoLine] : []), PromoAmount: promoAmount,
          UseMedicine: useMedicine ? 'TRUE' : 'FALSE', MedicineJSON: JSON.stringify(medicineItemsSeed), MedicineAmount: medicineAmount,
          UseOther: otherAmount > 0 ? 'TRUE' : 'FALSE', OtherItemsJSON: otherAmount > 0 ? JSON.stringify([{ label: 'อื่นๆ', amount: otherAmount }]) : '[]',
          // Round 20 cleanup: was hardcoded to `otherAmount` alone (harmless
          // only because medicineAmount was always 0 before this fix) — the
          // real formula (see saveRecord/CONFIG.HEADERS.Records doc comment)
          // is MedicineAmount + OtherAmount.
          SupplierPayAmount: round2_(medicineAmount + otherAmount)
        });
        if (status !== 'cancelled') seededGroups[doctor.id + '|' + dateStr_(recDate)] = { doctorId: doctor.id, date: dateStr_(recDate) };
        i++;
      }
    });
    Object.keys(seededGroups).forEach(function (key) {
      var g = seededGroups[key];
      recomputeDoctorDayShares_(g.doctorId, g.date);
    });

    // Expenses
    var expSheet = getSheet_(CONFIG.SHEETS.EXPENSES);
    for (var e = 0; e < 15; e++) {
      var cat = catIds.expense[e % catIds.expense.length];
      var d = new Date(now.getFullYear(), now.getMonth(), 1);
      d.setDate(d.getDate() - e * 6);
      appendRow_(expSheet, CONFIG.HEADERS.Expenses, {
        ExpID: 'EX' + nextSeq_('EXPENSE', 5), Date: dateStr_(d), Name: cat.name + ' ประจำเดือน', Category: cat.name,
        Amount: 2000 + (e % 5) * 1500, CreatedBy: 'admin', CreatedAt: nowStr_()
      });
    }

    // Orders + OrderLines (purchase orders)
    var orderSheet = getSheet_(CONFIG.SHEETS.ORDERS);
    var lineSheet = getSheet_(CONFIG.SHEETS.ORDER_LINES);
    var suppliers = ['บริษัท เมดิคอล ซัพพลาย จำกัด', 'ร้านเวชภัณฑ์ไทย', 'บริษัท บิวตี้ อิมปอร์ต จำกัด'];
    for (var o = 0; o < 5; o++) {
      var poId = 'PO' + nextSeq_('ORDER', 5);
      var d2 = new Date(now.getFullYear(), now.getMonth(), 1);
      d2.setDate(d2.getDate() - o * 15);
      var lineCount = 2 + (o % 3);
      var total = 0;
      var linesToWrite = [];
      for (var l = 0; l < lineCount; l++) {
        var item = stockIds[(o + l) % stockIds.length];
        var qty = 5 + l * 2;
        var unitPrice = 300 + l * 120;
        total += qty * unitPrice;
        linesToWrite.push({ LineID: poId + '-' + (l + 1), PoID: poId, Category: item.category, ItemName: item.name, Qty: qty, UnitPrice: unitPrice, LineTotal: qty * unitPrice });
      }
      appendRow_(orderSheet, CONFIG.HEADERS.Orders, { PoID: poId, Date: dateStr_(d2), Supplier: suppliers[o % suppliers.length], TotalAmount: total, CreatedBy: 'admin', CreatedAt: nowStr_() });
      linesToWrite.forEach(function (row) { appendRow_(lineSheet, CONFIG.HEADERS.OrderLines, row); });
    }

    // Clinic settings
    var settingsSheet = getSheet_(CONFIG.SHEETS.CLINIC_SETTINGS);
    var defaults = {
      name: 'VALIN CLINIC', address: '123 ถนนสุขุมวิท แขวงคลองตัน เขตคลองเตย กรุงเทพฯ 10110',
      phone: '02-123-4567', taxId: '0-1055-12345-67-8', email: 'contact@valinclinic.com', logoFileId: ''
    };
    Object.keys(defaults).forEach(function (key) {
      appendRow_(settingsSheet, CONFIG.HEADERS.ClinicSettings, { Key: key, Value: defaults[key] });
    });

    invalidateMasterCache_();
    invalidatePatientsLiteCache_();
    invalidateStockLiteCache_();
    return { ok: true, message: 'สร้างข้อมูลตัวอย่างเรียบร้อยแล้ว' };
  });
}

/* ======================================================================
   LINE — daily sales summary via the LINE Messaging API (push message).
   LINE Notify was permanently shut down 2025-03-31, so this uses a LINE
   Official Account's Channel Access Token (Script Property
   LINE_CHANNEL_ACCESS_TOKEN) to multicast a push. The message layout
   (Line.html, a Flex "bubble" card) and the plain-text fallback below
   both mirror the exact sectioned layout of the Records-calendar
   day-summary panel (see computeDailySummaryDetail_/dayPanelTemplate) —
   same numbers, same grouping, per the clinic's confirmed final design.

   Round 21: WHO the push goes to changed. It used to go to one single
   fixed id (a Script Property, LINE_TARGET_ID — usually a shared LINE
   group). It now goes to every admin-level Staff member's own LINE
   account instead — see getLineAdminTargets_() immediately below — per
   an explicit client request to restrict this to admins with a
   registered LineID, rather than a shared group anyone could be added
   to. This reuses the exact same Staff.LineID field the Round 20
   webhook already authorizes chat commands against, just filtered
   further down to Level === 'admin' for this specific recipient list
   (the webhook's own chat-command authorization is unchanged — it still
   allows any Active staff member, not just admins; see assumption #31 in
   the project notes for that still-open question).
   ====================================================================== */

/** Round 22: every Staff row this round's "ส่งข้อมูลไป LINE" push should
 *  target — admin-level (`Level === 'admin'`), Active, and with a
 *  non-empty LineID. Returns `{lineId, name}` objects (Round 24: added
 *  `name` so the caller can tell the user WHO the push actually went to,
 *  not just how many — see sendLineDailySummary below). Deliberately
 *  mirrors findLineAuthorizedStaff_()'s Active-row check below for
 *  consistency (an inactive admin's old LineID shouldn't keep receiving
 *  pushes). */
function getLineAdminTargets_() {
  return getRows_(getSheet_(CONFIG.SHEETS.STAFF)).filter(function (r) {
    return r.Level === 'admin' &&
      r.LineID && String(r.LineID).trim() &&
      (r.Active === true || String(r.Active).toUpperCase() === 'TRUE');
  }).map(function (r) { return { lineId: String(r.LineID).trim(), name: r.Name || '(ไม่มีชื่อ)' }; });
}

function moneyFmt_(n) {
  return Number(n || 0).toLocaleString('en-US', { minimumFractionDigits: 2, maximumFractionDigits: 2 });
}

function getClinicName_() {
  var row = getRows_(getSheet_(CONFIG.SHEETS.CLINIC_SETTINGS)).filter(function (r) { return r.Key === 'name'; })[0];
  return row && row.Value ? row.Value : 'VALIN CLINIC';
}

// Round 18 item 1: mirrors the day panel's new flat, numbered 8-row shape
// (dayPanelTemplate/daySummaryItem in client.js) — 1.1 ยอดชำระทั้งหมด, 1.2
// จำนวนเคส, 1.3 ส่วนแบ่งแพทย์, 1.4 ส่วนแบ่งเซลล์, 2 canula, 3 โปร, 4
// รายการอื่นๆ, 5 ยอดสุทธิ (now the same totalCollected total broken down by
// recipient instead of by payment method — see the client-side comment for
// the identity this relies on). The plain-text fallback has no collapse/
// expand concept, so every group's detail is always printed inline — same
// as before this round, just reorganized/relabeled to match.
function buildPlainSummaryText_(s, clinicName) {
  var lines = [];
  lines.push((clinicName || 'VALIN CLINIC') + ' — สรุปยอดประจำวัน');
  lines.push(s.date);
  lines.push('');
  lines.push('ยอดชำระทั้งหมด: ' + moneyFmt_(s.totalCollected) + ' บาท');
  lines.push('- เงินสด: ' + moneyFmt_(s.cashAmount) + ' บาท');
  lines.push('- เงินโอน: ' + moneyFmt_(s.transferAmount) + ' บาท');
  if (s.cardAmount) lines.push('- บัตรเครดิต: ' + moneyFmt_(s.cardAmount) + ' บาท');
  lines.push('');
  lines.push('จำนวนเคส: ' + s.patientCount + ' คน');
  lines.push('');
  lines.push('ส่วนแบ่งแพทย์ ทั้งหมด: ' + moneyFmt_(s.doctorTotal) + ' บาท');
  if (!s.doctorShares.length) lines.push('ไม่มีรายการ');
  s.doctorShares.forEach(function (d) { lines.push(d.name + (d.topped ? ' *' : '') + ': ' + moneyFmt_(d.amount) + ' บาท'); });
  if (s.hasToppedDoctor) lines.push('* ได้รับตามยอดขั้นต่ำที่คลินิกรับประกัน');
  lines.push('');
  lines.push('ส่วนแบ่งเซลล์ ทั้งหมด: ' + moneyFmt_(s.salesTotal) + ' บาท');
  if (!s.salesShares.length) lines.push('ไม่มีรายการ');
  s.salesShares.forEach(function (d) { lines.push(d.name + ': ' + moneyFmt_(d.amount) + ' บาท'); });
  lines.push('');
  lines.push('ยอด canula (เข็มทู่): ' + moneyFmt_(s.canulaTotal) + ' บาท');
  lines.push('');
  lines.push('ยอดรายการโปรหน้าคลีนิค: ' + moneyFmt_(s.promoAmountTotal) + ' บาท');
  (s.promoItems || []).forEach(function (p) { lines.push(p.name + ' (x' + p.qty + '): ' + moneyFmt_(p.amount) + ' บาท'); });
  lines.push('');
  lines.push('รายการอื่นๆ: ' + moneyFmt_(s.supplierPayTotal) + ' บาท');
  if (!s.otherItems.length && !(s.medicineItems && s.medicineItems.length)) lines.push('ไม่มีรายการ');
  // Round 15 item 3: "ชื่อรายการ จำนวนเงิน" — item name leads, amount
  // follows (was reversed, and only showed the patient name with no item
  // name at all).
  (s.medicineItems || []).forEach(function (m) { lines.push((m.name || 'ยา') + ' (' + m.patientName + '): ' + moneyFmt_(m.amount) + ' บาท'); });
  s.otherItems.forEach(function (o) { lines.push((o.name || 'อื่นๆ') + ' (' + o.patientName + '): ' + moneyFmt_(o.amount) + ' บาท'); });
  lines.push('');
  // Round 19 item 1: ยอดสุทธิ = จำนวนเงินทั้งหมด - ส่วนแบ่งแพทย์ - ส่วนแบ่งเซลล์
  // - ส่วนแบ่ง Supplier (== s.clinicNet, per the revenue-split identity) —
  // reverts the Round 18 "totalCollected reframed by recipient" reading.
  lines.push('ยอดสุทธิ: ' + moneyFmt_(s.clinicNet) + ' บาท');
  lines.push('จำนวนเงินทั้งหมด: ' + moneyFmt_(s.totalCollected) + ' บาท');
  lines.push('- ส่วนแบ่งแพทย์: ' + moneyFmt_(s.doctorTotal) + ' บาท');
  lines.push('- ส่วนแบ่งเซลล์: ' + moneyFmt_(s.salesTotal) + ' บาท');
  lines.push('- ส่วนแบ่ง Supplier: ' + moneyFmt_(s.supplierPayTotal) + ' บาท');
  return lines.join('\n');
}

/** Builds the one Flex Message object (or its plain-text fallback) for a
 *  day's summary — shared by the push button (sendLineDailySummary), the
 *  daily time-based trigger, and (Round 20 item 2) the new webhook chat
 *  commands below, so every surface that sends a summary to LINE renders it
 *  through the exact same code path. */
function buildLineSummaryMessage_(summary, clinicName) {
  try {
    var tpl = HtmlService.createTemplateFromFile('Line');
    tpl.data = summary;
    tpl.clinicName = clinicName;
    return JSON.parse(tpl.evaluate().getContent());
  } catch (e) {
    return { type: 'text', text: buildPlainSummaryText_(summary, clinicName) }; // fallback if Line.html ever fails to evaluate
  }
}

/** token may be omitted when called from a time-based trigger (no logged-in user).
 *  Round 22: pushes to every admin-level, Active Staff row's own LineID
 *  (getLineAdminTargets_()) instead of one fixed Script Property id — see
 *  the LINE section's block comment above for why. Round 24: the returned
 *  success message now also names WHO it was sent to, not just how many —
 *  so the "ส่งข้อมูลไป LINE" button's toast lets the clicker immediately
 *  confirm the right admins received it. */
function sendLineDailySummary(token, dateStr) {
  return safeCall_(function () {
    var session = token ? requireAuth_(token, 'staff') : null;
    var props = PropertiesService.getScriptProperties();
    var channelToken = props.getProperty(CONFIG.PROPS.LINE_CHANNEL_ACCESS_TOKEN);
    if (!channelToken) throw new Error('ยังไม่ได้ตั้งค่า LINE_CHANNEL_ACCESS_TOKEN ใน Script Properties');

    var targets = getLineAdminTargets_();
    if (!targets.length) throw new Error('ยังไม่มีผู้ดูแลระบบ (Level = admin) ที่กรอก LINE ID ไว้ในหน้าตั้งค่าพนักงาน — กรุณาให้แอดมินอย่างน้อย 1 คนพิมพ์ "Line ID ของฉัน" ในแชท LINE เพื่อรับรหัสของตนเอง แล้วนำไปกรอกในช่อง LINE ID ของบัญชีนั้น');
    var targetIds = targets.map(function (t) { return t.lineId; });
    var targetNames = targets.map(function (t) { return t.name; });

    var date = dateStr || todayStr_();
    var summary = computeDailySummaryDetail_(date);
    var clinicName = getClinicName_();
    var message = buildLineSummaryMessage_(summary, clinicName);

    // LINE's multicast endpoint accepts up to 500 user ids in a single
    // call, so N registered admins costs 1 API call, not N pushes.
    var resp = UrlFetchApp.fetch('https://api.line.me/v2/bot/message/multicast', {
      method: 'post', contentType: 'application/json',
      headers: { Authorization: 'Bearer ' + channelToken },
      payload: JSON.stringify({ to: targetIds, messages: [message] }),
      muteHttpExceptions: true
    });
    if (resp.getResponseCode() >= 300) throw new Error('ส่งข้อความ LINE ไม่สำเร็จ: ' + resp.getContentText());
    writeAudit_(session, 'LINE_PUSH', 'Summary', date, 'to ' + targetIds.length + ' admin(s): ' + targetNames.join(', '));
    return {
      ok: true,
      message: 'ส่งสรุปยอดประจำวันไปยัง LINE เรียบร้อยแล้ว (' + targetIds.length + ' คน: ' + targetNames.join(', ') + ')',
      recipientCount: targetIds.length,
      recipientNames: targetNames
    };
  });
}

/** Wrapper with no params, for use as a time-based trigger handler. */
function sendLineDailySummaryTrigger_() {
  sendLineDailySummary(null, todayStr_());
}

/** Admin utility: install/replace a daily time-based trigger that sends the summary automatically. */
function installDailyLineSummaryTrigger(token, hour) {
  return safeCall_(function () {
    requireAuth_(token, 'admin');
    ScriptApp.getProjectTriggers().forEach(function (t) {
      if (t.getHandlerFunction() === 'sendLineDailySummaryTrigger_') ScriptApp.deleteTrigger(t);
    });
    ScriptApp.newTrigger('sendLineDailySummaryTrigger_').timeBased().atHour(hour || 20).everyDays(1).create();
    return { ok: true, message: 'ตั้งเวลาส่งสรุปยอด LINE อัตโนมัติทุกวัน เวลา ' + (hour || 20) + ':00 แล้ว' };
  });
}

/* ======================================================================
   LINE — interactive webhook chat commands (Round 20 item 2).

   The client had previously written a standalone LINE bot script (a
   doPost webhook that answered "สรุปยอดวันนี้"/"สรุปยอดวันอื่น"/"สรุปยอดวันที่
   yyyy-MM-dd"/"Line ID ของฉัน" typed in LINE) against an older, unrelated
   raw spreadsheet layout (its own COL mapping, a "2-1 รายการรักษา" sheet,
   its own from-scratch getDailySummary()/generateFlexMessage() — none of
   which exist in or match this app's schema). This section ports the
   COMMAND HANDLING (which text triggers which reply) and the calendar
   day-picker idea over to run on THIS app's own data engine instead:
   every command below calls computeDailySummaryDetail_/getRows_ against
   the Records sheet and renders through buildLineSummaryMessage_ (the
   exact same function the push button uses), so the web app, the daily
   auto-push, and this chat bot can never disagree with each other.

   AUTHORIZATION: LINE webhook requests carry no Apps Script session token
   (the person is chatting from LINE, not logged into the web app), so
   there's no requireAuth_() to call. Instead, the summary commands (not
   "Line ID ของฉัน", which anyone can use to look up their own id) require
   the sender's LINE userId to match an ACTIVE Staff row's LineID column —
   the exact same field already editable from ตั้งค่า > ตั้งค่าพนักงาน (see
   client.js's staff modal, "LINE ID" field). This reuses the app's own
   staff roster instead of asking the client to maintain a second,
   separate "who's allowed" list the way the reference script's (disabled)
   isLineAdmin()/"LineID" sheet did — an admin registers a staff member's
   id there, and that's the one and only place access is granted. To learn
   their own id, a staff member messages the bot "Line ID ของฉัน" first.

   SETUP: after deploying, a web app /exec URL must be registered as this
   LINE Official Account's Webhook URL (LINE Developers Console > Messaging
   API > Webhook settings > "Use webhook" ON) — see คู่มือติดตั้ง.md for the
   full walkthrough. No new Script Property is needed; this reuses the
   same LINE_CHANNEL_ACCESS_TOKEN already set up for the push button (a
   Channel Access Token is valid for both the push and reply Messaging API
   endpoints).

   **(Round 23) Use a SEPARATE Web app deployment's URL for this, distinct
   from the URL staff use to log in** — create a second deployment via
   Deploy > Manage deployments > New deployment (same settings: Web app,
   Execute as "Me", access "Anyone") and give LINE that one's /exec URL
   instead. Both deployments execute the identical doGet/doPost code (a
   single Apps Script project cannot behave differently per deployment
   URL), so this doesn't change what runs — it only gives LINE its own
   address, separate from the one handed out to clinic staff. This means
   any future code push needs BOTH deployments re-pointed at the new
   version (Manage deployments > pencil icon on each > "Version: New
   version" > Deploy), or one of the two URLs will keep serving old code.

   NOTE ON SIGNATURE VERIFICATION: Apps Script's doPost(e) does not expose
   the incoming request's HTTP headers, so the X-Line-Signature check LINE
   recommends cannot be implemented here — this is a known Apps Script
   platform limitation, not an oversight (the reference script had the
   same gap). The Staff-roster authorization above is the practical
   safeguard for the data-returning commands instead.
   ====================================================================== */

/** Web app entry point for LINE's webhook. Always responds 200 (LINE's
 *  "Verify" button and any retry both just need a 2xx, not a particular
 *  body), and never lets one malformed/unexpected event stop the others in
 *  the same batch from being handled. */
function doPost(e) {
  try {
    if (!e || !e.postData || !e.postData.contents) return ContentService.createTextOutput('OK');
    var json = JSON.parse(e.postData.contents);
    var events = json.events || [];
    events.forEach(function (event) {
      try { handleLineWebhookEvent_(event); } catch (err) { console.error('handleLineWebhookEvent_ error: ' + err); }
    });
  } catch (err) {
    console.error('doPost error: ' + err);
  }
  return ContentService.createTextOutput('OK');
}

/** Dispatches one LINE webhook event to the matching chat command, or does
 *  nothing for anything else (stickers, images, unrecognized text, etc.) —
 *  same "stay quiet unless it's a known command" behavior as the reference
 *  script, important for a bot that may sit in a shared group chat. */
function handleLineWebhookEvent_(event) {
  if (!event || event.type !== 'message' || !event.message || event.message.type !== 'text') return;
  var replyToken = event.replyToken;
  var userId = event.source && event.source.userId;
  var text = String(event.message.text || '').trim();

  if (text === 'Line ID ของฉัน') {
    replyLineMessage_(replyToken, { type: 'text', text: 'LINE ID ของคุณคือ\n' + (userId || '(ไม่พบ)') });
    return;
  }

  var isTodayCmd = text === 'สรุปยอดวันนี้';
  var isOtherDayCmd = text === 'สรุปยอดวันอื่น' || text.indexOf('สรุปยอดวันอื่น ') === 0;
  var isSpecificDateCmd = /^สรุปยอดวันที่ \d{4}-\d{2}-\d{2}$/.test(text);
  if (!isTodayCmd && !isOtherDayCmd && !isSpecificDateCmd) return;

  var staff = findLineAuthorizedStaff_(userId);
  if (!staff) {
    replyLineMessage_(replyToken, { type: 'text', text: '⛔ LINE ID นี้ยังไม่ได้รับสิทธิ์ใช้งานคำสั่งนี้\nกรุณาแจ้งผู้ดูแลระบบให้เพิ่ม LINE ID นี้ในข้อมูลพนักงาน (ตั้งค่า > ตั้งค่าพนักงาน)' });
    return;
  }

  if (isTodayCmd) {
    replyLineDailySummary_(replyToken, todayStr_());
  } else if (isSpecificDateCmd) {
    replyLineDailySummary_(replyToken, text.replace('สรุปยอดวันที่', '').trim());
  } else {
    // "สรุปยอดวันอื่น" or "สรุปยอดวันอื่น yyyy-MM" -> show a tappable calendar
    // for that month (current month if none given); tapping a day sends
    // "สรุปยอดวันที่ yyyy-MM-dd" right back through this same webhook.
    var parts = text.split(' ');
    var tz = Session.getScriptTimeZone() || 'Asia/Bangkok';
    var year, month;
    if (parts.length >= 2 && /^\d{4}-\d{2}$/.test(parts[1])) {
      var ym = parts[1].split('-');
      year = parseInt(ym[0], 10); month = parseInt(ym[1], 10);
    } else {
      var now = new Date();
      year = parseInt(Utilities.formatDate(now, tz, 'yyyy'), 10);
      month = parseInt(Utilities.formatDate(now, tz, 'M'), 10);
    }
    replyLineMessage_(replyToken, buildLineCalendarMessage_(year, month));
  }
}

function replyLineDailySummary_(replyToken, dateStr) {
  var summary = computeDailySummaryDetail_(dateStr);
  var clinicName = getClinicName_();
  replyLineMessage_(replyToken, buildLineSummaryMessage_(summary, clinicName));
}

function replyLineMessage_(replyToken, message) {
  if (!replyToken) return;
  var channelToken = PropertiesService.getScriptProperties().getProperty(CONFIG.PROPS.LINE_CHANNEL_ACCESS_TOKEN);
  if (!channelToken) return; // not configured yet — silently no-op, same as any other misconfigured-integration guard in this file
  UrlFetchApp.fetch('https://api.line.me/v2/bot/message/reply', {
    method: 'post', contentType: 'application/json',
    headers: { Authorization: 'Bearer ' + channelToken },
    payload: JSON.stringify({ replyToken: replyToken, messages: [message] }),
    muteHttpExceptions: true
  });
}

/** Looks up the Staff row (if any) whose LineID column matches the given
 *  LINE userId AND is marked Active — the same "Active" flag that already
 *  gates that staff member's own web-app login. Returns null (unauthorized)
 *  for an empty/unregistered userId. */
function findLineAuthorizedStaff_(userId) {
  if (!userId) return null;
  var rows = getRows_(getSheet_(CONFIG.SHEETS.STAFF));
  var match = rows.filter(function (r) {
    return r.LineID && String(r.LineID).trim() === String(userId).trim() &&
      (r.Active === true || String(r.Active).toUpperCase() === 'TRUE');
  })[0];
  return match || null;
}

/** Which days in `year`-`month` have at least one paid Records row —
 *  powers the "●" markers on the webhook's calendar day-picker. */
function getDaysWithDataInMonth_(year, month) {
  var prefix = year + '-' + ('0' + month).slice(-2);
  var days = {};
  getRows_(getSheet_(CONFIG.SHEETS.RECORDS)).forEach(function (r) {
    if (r.Status !== 'paid') return;
    var ymd = dateStr_(r.Date);
    if (ymd && ymd.indexOf(prefix) === 0) days[Number(ymd.substring(8, 10))] = true;
  });
  return days;
}

/** A tappable month calendar Flex Message — ported from the reference
 *  script's generateCalendarFlexMessage, restyled to this app's own navy
 *  header (matching buildLineSummaryMessage_'s bubble) instead of the
 *  reference script's unrelated header color. Tapping a day fires a
 *  "message" action with the text "สรุปยอดวันที่ yyyy-MM-dd", which comes
 *  straight back through doPost/handleLineWebhookEvent_ like any other
 *  typed command. */
function buildLineCalendarMessage_(year, month) {
  var daysWithData = getDaysWithDataInMonth_(year, month);
  var daysInMonth = new Date(year, month, 0).getDate();
  var firstWeekday = new Date(year, month - 1, 1).getDay();
  var weekdayLabels = ['อา', 'จ', 'อ', 'พ', 'พฤ', 'ศ', 'ส'];
  var monthNamesTh = ['มกราคม', 'กุมภาพันธ์', 'มีนาคม', 'เมษายน', 'พฤษภาคม', 'มิถุนายน', 'กรกฎาคม', 'สิงหาคม', 'กันยายน', 'ตุลาคม', 'พฤศจิกายน', 'ธันวาคม'];

  var headerRow = {
    type: 'box', layout: 'horizontal',
    contents: weekdayLabels.map(function (d) { return { type: 'text', text: d, size: 'xs', align: 'center', color: '#8891AC', weight: 'bold', flex: 1 }; })
  };

  var weekRows = [headerRow];
  var totalCells = firstWeekday + daysInMonth;
  var totalWeeks = Math.ceil(totalCells / 7);
  var dayCounter = 1;

  for (var w = 0; w < totalWeeks; w++) {
    var cells = [];
    for (var d = 0; d < 7; d++) {
      var cellIndex = w * 7 + d;
      if (cellIndex < firstWeekday || dayCounter > daysInMonth) {
        cells.push({ type: 'box', layout: 'vertical', flex: 1, contents: [{ type: 'text', text: ' ', size: 'xs' }] });
      } else {
        var day = dayCounter;
        var hasData = !!daysWithData[day];
        var dateStrForDay = year + '-' + ('0' + month).slice(-2) + '-' + ('0' + day).slice(-2);
        cells.push({
          type: 'box', layout: 'vertical', flex: 1, alignItems: 'center', spacing: 'xs', paddingAll: '4px', cornerRadius: 'sm',
          action: { type: 'message', label: 'เลือกวันที่ ' + dateStrForDay, text: 'สรุปยอดวันที่ ' + dateStrForDay },
          contents: [
            { type: 'text', text: String(day), size: 'sm', align: 'center', weight: hasData ? 'bold' : 'regular', color: hasData ? '#1E3A72' : '#1C2440' },
            { type: 'text', text: hasData ? '●' : ' ', size: 'xxs', align: 'center', color: '#AD8C46' }
          ]
        });
        dayCounter++;
      }
    }
    weekRows.push({ type: 'box', layout: 'horizontal', spacing: 'xs', margin: 'sm', contents: cells });
  }

  return {
    type: 'flex',
    altText: 'เลือกวันที่เพื่อดูสรุปยอด - ' + monthNamesTh[month - 1] + ' ' + (year + 543),
    contents: {
      type: 'bubble', size: 'mega',
      header: {
        type: 'box', layout: 'vertical', backgroundColor: '#1E3A72', paddingAll: '16px',
        contents: [
          { type: 'text', text: 'เลือกวันที่ต้องการสรุปยอด', color: '#FFFFFF', weight: 'bold', size: 'md' },
          { type: 'text', text: monthNamesTh[month - 1] + ' ' + (year + 543), color: '#E2E6EE', size: 'sm', margin: 'sm' }
        ]
      },
      body: { type: 'box', layout: 'vertical', spacing: 'sm', paddingAll: '12px', contents: weekRows },
      footer: {
        type: 'box', layout: 'vertical', paddingAll: '12px',
        contents: [{ type: 'text', text: '● = วันที่มียอดขาย  •  แตะวันที่เพื่อดูสรุป', size: 'xs', color: '#8891AC', align: 'center' }]
      }
    }
  };
}