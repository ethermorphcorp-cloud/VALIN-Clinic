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
 *                  PhotoFileId, JoinedDate, CreatedAt, UpdatedAt, Nickname, HN
 *   Records       Core: RID, Date, PatientID, PatientName, Phone,
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
 *   Procedures     ProcID, Name, Category, Price, Status, CreatedAt, UpdatedAt
 *                  — รายการหัตถการ picked by the record modal (ProcedureJSON).
 *   MedDeposits / MedWithdrawals — ยาฝาก lots and จัดยา-เบิกยา lines, see
 *                  CONFIG.HEADERS.
 *   Categories     CategoryID, Type, Name, SortOrder      (Type: treatment (หัตถการ)|stock|expense)
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
 *   2. (Optional) Run `seedDemoData` to replace the business data with
 *      sample data in the current format, or `clearAllData` to empty it.
 *      Both keep the Staff sheet and ClinicSettings (company info/logo).
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
    MED_WITHDRAWALS: 'MedWithdrawals',
    PROCEDURES: 'Procedures',
    AUDIT_LOG: 'AuditLog',
    DAILY_SUMMARY: 'DailySummary',
    BACKUPS: 'Backups'
  },
  HEADERS: {
    // NOTE ON SCHEMA CHANGES: new columns are always appended at the END of
    // these arrays, never inserted/renamed in the middle — ensureSheet_()
    // rewrites row 1 (headers only, never data) whenever this array doesn't
    // match what's on the sheet, so appending keeps every existing row's
    // values correctly aligned under their original columns. Re-running
    // setupProject() after a schema change like this is always safe.
    // HN = the clinic's own patient number, typed by staff (optional, unique
    // when filled) — separate from the auto-generated ID.
    Patients: ['ID', 'FirstName', 'LastName', 'Phone', 'DOB', 'Gender', 'Email', 'PhotoFileId', 'JoinedDate', 'CreatedAt', 'UpdatedAt', 'Nickname', 'HN'],
    // Revenue-split fields (added for the commission engine): OtherAmount is
    // now the SUM of the round-13 repeatable "รายการอื่นๆ" lines (see
    // OtherItemsJSON) — still excluded from anyone's commission share, only
    // feeds SupplierPayAmount. SalesCommissionPct/DoctorCommissionPct
    // snapshot the rate actually used at save time (so later rate changes on
    // the Doctors/Sales sheets don't retroactively rewrite historical
    // splits). DoctorShareRaw is the plain commission BEFORE the
    // daily-minimum top-up — as of round 13 it is
    // Amount*DoctorCommissionPct/100 + sum(each PromoJSON line's amount *
    // that promo's own DoctorSharePct/100) (see buildDoctorShares_).
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
    // UseProcedure/ProcedureJSON: รายการหัตถการ lines picked from the
    // Procedures sheet, each {procId,name,price,amount,doctorId}; Amount
    // (above) is their sum on records saved since. Older records only have
    // Amount.
    // DoctorSharesJSON: one entry per doctor on the record ({doctorId, name,
    // pct, raw, share, note}) — each หัตถการ/โปร line names its own doctor, so
    // one visit can pay several doctors. DoctorID/DoctorName/DoctorShareRaw/
    // DoctorShare/DoctorShareNote keep the first doctor / the names joined /
    // the totals for older readers (see applyDoctorShares_). Records from
    // before it have one doctor, read from those columns (recordDoctorShares_).
    // WithdrawJSON: the จัดยา-เบิกยา lines this record took in, each
    // {withdrawId, itemId, name, qtyDispensed, qtyUsed, qtyDeposit, expiry} —
    // not charged here (the sales person sells them); only their ยาฝาก part
    // matters to the record.
    // VisitStatus: 'dispensed' จัดยาแล้ว → 'waiting_doctor' รอพบแพทย์ →
    // 'with_doctor' พบแพทย์ → 'waiting_payment' รอชำระเงิน → 'completed' เสร็จสิ้น (= paid); older
    // ''/'scheduled' read as 'dispensed', a stored 'completed' that isn't paid
    // reads as 'waiting_payment' (see visitStatusOf_). QueueDate/QueueDoctor:
    // the day's queue number, taken when the customer arrives (รอพบแพทย์) and
    // kept through the later steps; QueuePay is no longer used. VisitDoctorID/
    // VisitDoctorName: the doctor chosen when it moved to พบแพทย์.
    // ClinicShare is now everything collected (incl. ยาเบิกเพิ่ม and
    // รายการอื่นๆ) minus the sales and doctor shares — there is no separate
    // supplier share any more (SupplierPayAmount is still written, unused).
    Records: ['RID', 'Date', 'PatientID', 'PatientName', 'Phone', 'ServiceCategory', 'Detail', 'DoctorID', 'DoctorName', 'SalesID', 'SalesName', 'PaymentMethod', 'SuppliesJSON', 'UseCanula', 'CanulaAmount', 'Amount', 'Status', 'ReceiptNo', 'Note', 'CreatedBy', 'CreatedAt', 'UpdatedAt', 'OtherAmount', 'SalesCommissionPct', 'SalesShare', 'DoctorCommissionPct', 'DoctorShareRaw', 'DoctorShare', 'DoctorShareNote', 'ClinicShare', 'VisitStatus', 'UsePromo', 'PromoJSON', 'PromoAmount', 'UseMedicine', 'MedicineJSON', 'MedicineAmount', 'UseOther', 'OtherItemsJSON', 'SupplierPayAmount', 'OpdImageFileId', 'UseSupplies', 'UseOwnMed', 'OwnMedJSON', 'UseDepositUse', 'DepositUseJSON', 'UseMessenger', 'MessengerJSON', 'StockApplied', 'UseProcedure', 'ProcedureJSON', 'DoctorSharesJSON', 'WithdrawJSON', 'QueueDate', 'QueueDoctor', 'QueuePay', 'VisitDoctorID', 'VisitDoctorName', 'CanulaJSON', 'UseOralMed', 'OralMedJSON', 'OralMedAmount'],
    // UseOralMed/OralMedJSON/OralMedAmount: รายการจ่ายยาทาน — lines picked from the รายการยาทาน stock (StockType 'oral'), each
    // {itemId, name, qty, unit, unitPrice, amount = qty × unitPrice}; charged in the ยอดรับชำระ (recordCollected_) and taken
    // out of Stock (recordStockEffects_). OralMedAmount is their sum.
    // CanulaJSON: the รายการ Canula (เข็มทู่) lines, each {name, qty, amount}; CanulaAmount is their sum (records from before
    // keep just CanulaAmount and read as one line). MedicineJSON holds BOTH ยา groups of the record form: lines with
    // charged=true are รายการซื้อยาเพิ่มหน้าคลีนิค (คิดค่ายา) — {…, unitPrice, amount = qtyDispensed × unitPrice, unitCost} —
    // the others are (ไม่คิดค่ายา) with amount 0; see medicineLines_ / splitMedicineLines_. ProcedureJSON lines also carry
    // a free-text detail.
    // MedDeposits — ยาฝาก, one row per lot: a record's medicine line that is
    // not all used creates a lot (RecordID = that record) with its own
    // ExpiryDate; รายการใช้ยาฝาก takes from the patient's lots, earliest
    // expiry first (see planDeposits_). QtyDeposited is what the lot started
    // with. Rows from before lots existed are one balance per patient + item
    // with blank RecordID/ExpiryDate/QtyDeposited.
    MedDeposits: ['DepositID', 'Date', 'PatientID', 'PatientName', 'ItemName', 'QtyRemaining', 'CreatedBy', 'ItemID', 'UpdatedAt', 'ExpiryDate', 'RecordID', 'QtyDeposited'],
    // MedWithdrawals — เมนู "จัดยา-เบิกยา": one row per medicine line. UnitCost
    // is the item's ราคาทุน snapshotted at save time (Amount = Qty * UnitCost);
    // both are admin-only on the way out. Qty is taken out of Stock.
    // BatchID groups the lines saved together in one save. UnitPrice/
    // UnitPrice/SaleAmount are no longer used (the customer pays the sales
    // person for the ยา directly, so no selling price is kept; the columns stay
    // for older rows); RecordID = the บันทึกการรักษา the line belongs to.
    MedWithdrawals: ['WithdrawID', 'Date', 'PatientID', 'PatientName', 'Phone', 'SalesID', 'SalesName', 'ItemID', 'ItemName', 'Qty', 'UnitCost', 'Amount', 'Note', 'CreatedBy', 'CreatedAt', 'BatchID', 'UnitPrice', 'SaleAmount', 'RecordID', 'PhotoFileIds'],
    // Procedures — รายการหัตถการ (จัดการสินค้า/บริการ). Category is a
    // Categories row of Type 'treatment'.
    Procedures: ['ProcID', 'Name', 'Category', 'Price', 'Status', 'CreatedAt', 'UpdatedAt'],
    // StockType: 'supply' (เวชภัณฑ์) | 'medicine' (ยา) | 'oral' (ยาทาน, one ราคาขาย kept in Price1) — splits the Stock page
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
    // StockApplied/ItemID appended later: lines picked from the stock list
    // carry ItemID and their Qty is added to Stock — only on POs created
    // since then (StockApplied=TRUE), so editing an older PO never re-adds it.
    // OrderType: 'sales' = ยา(จากเซลล์), Supplier holds the sales person's name; 'other' (or blank) = อื่นๆ, Supplier is the vendor.
    Orders: ['PoID', 'Date', 'Supplier', 'TotalAmount', 'CreatedBy', 'CreatedAt', 'StockApplied', 'OrderType'],
    OrderLines: ['LineID', 'PoID', 'Category', 'ItemName', 'Qty', 'UnitPrice', 'LineTotal', 'ItemID'],
    Doctors: ['DoctorID', 'Name', 'Status', 'MinAmount', 'CreatedAt', 'CommissionPct'],
    Sales: ['SalesID', 'Name', 'Status', 'CommissionRate', 'MinAmount', 'CreatedAt'],
    // Promotions (round 12, item 3) — reference/display table for the
    // "จัดการโปรโมชันหน้าคลินิก" page. NOTE: round 12 deliberately kept this
    // standalone from the treatment record modal / revenue-split engine; round
    // 13 explicitly REVERSES that — the record modal's "รายการโปร" picker now
    // reads this list live, and each promo's DoctorSharePct feeds directly
    // into ส่วนแบ่งแพทย์ (see buildDoctorShares_ / Records.PromoJSON above).
    // ClinicSharePct is always derived as (100 - DoctorSharePct), never
    // stored, so it can never drift.
    Promotions: ['PromoID', 'Name', 'Price', 'DoctorSharePct', 'Status', 'CreatedAt', 'UpdatedAt'],
    Categories: ['CategoryID', 'Type', 'Name', 'SortOrder'],
    Staff: ['StaffID', 'Name', 'Role', 'Username', 'PasswordHash', 'PasswordSalt', 'LineID', 'Level', 'Active', 'CreatedAt'],
    ClinicSettings: ['Key', 'Value'],
    Counters: ['Key', 'LastSeq'],
    AuditLog: ['Timestamp', 'User', 'Action', 'Entity', 'EntityID', 'Detail'],
    // DailySummary — what the dashboard still needs from data that was moved to a backup file (see backupData): one row per
    // Date + Metric (+ Key), long format. Metrics: customerRevenue, proc, promo, canula, drug (money collected from customers,
    // paid records), salesIncome (รายงานรายเซลล์ ราคาทุน), doctorShareTotal, salesShareTotal, doctorShare[Key=name],
    // salesShare[Key=name], promoLine[Key=name], expense[Key=category], orders, patient[Key=PatientID, Count=1].
    DailySummary: ['Date', 'Metric', 'Key', 'Amount', 'Count', 'BackupID'],
    // Backups — one row per backup file: where it is and what it holds. Status 'เสร็จสิ้น' once the main sheets were cleaned.
    Backups: ['BackupID', 'CreatedAt', 'CreatedBy', 'FromDate', 'ToDate', 'CutoffDate', 'FileName', 'SpreadsheetID', 'SpreadsheetUrl', 'FolderID', 'CountsJSON', 'Status']
  },
  // Columns that must be stored/read as literal text so leading zeros
  // (phone numbers, ID codes) are never lost to number coercion.
  TEXT_COLUMNS: {
    Patients: ['Phone', 'HN'],
    Records: ['Phone', 'ReceiptNo'],
    MedWithdrawals: ['Phone'],
    Staff: ['Username']
  },
  // Numeric columns whose cell format is reset to automatic on every write
  // (see applyTextColumnFormat_), so a cell left date-formatted from earlier
  // data can't turn a quantity or price into a date.
  NUMBER_COLUMNS: {
    Stock: ['Qty', 'MinQty', 'CostPrice', 'Price1', 'Price3', 'Price5', 'Price10', 'Price20', 'Price50'],
    DailySummary: ['Amount', 'Count'],
    StockLog: ['QtyIn', 'QtyRemaining', 'QtyOut'],
    MedDeposits: ['QtyRemaining', 'QtyDeposited'],
    MedWithdrawals: ['Qty', 'UnitCost', 'Amount', 'UnitPrice', 'SaleAmount'],
    Procedures: ['Price']
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
    MASTER_DATA_KEY: 'MASTER_DATA_V2', // V2: carries `procedures`
    MASTER_DATA_TTL: 21600, // 6h, the max CacheService allows
    SESSION_PREFIX: 'SESSION_',
    SESSION_TTL: 21600,     // 6h
    RECEIPT_TICKET_PREFIX: 'RTKT_',
    RECEIPT_TICKET_TTL: 300, // 5 minutes
    // Lightweight full-list caches used to power client-side autocomplete
    // (patient/stock suggest boxes) without a server round-trip per keystroke.
    PATIENTS_LITE_KEY: 'PATIENTS_LITE_V2', // V2: entries carry `hn`
    STOCK_LITE_KEY: 'STOCK_LITE_V4', // V4: entries also carry `qty` (พร้อมใช้, for the pickers); never cost
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
  PAGE_SIZE_DEFAULT: 15
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
/*
 * Write-scope cache. A save reads the same sheets several times (Stock three times, Doctors three, Counters once per id...) and
 * every read is a full Sheets API call. Inside beginWriteScope_() .. endWriteScope_() — which a write operation opens only
 * AFTER it holds the script lock, so nobody else can change the sheets meanwhile — each sheet is read once; the write
 * helpers below keep the cached copy in step with what they write. Outside a scope nothing is cached.
 */
var _rowsCache_ = null;
function beginWriteScope_() { _rowsCache_ = {}; }
function endWriteScope_() { _rowsCache_ = null; }
function sheetValues_(sheet) {
  if (!_rowsCache_) return sheet.getDataRange().getValues();
  var name = sheet.getName();
  return _rowsCache_[name] || (_rowsCache_[name] = sheet.getDataRange().getValues());
}
function cachePutRow_(sheet, rowIndex, vals) {
  if (!_rowsCache_) return;
  var cached = _rowsCache_[sheet.getName()];
  if (!cached) return;
  while (cached.length < rowIndex - 1) cached.push([]);
  cached[rowIndex - 1] = vals;
}

function getRows_(sheet) {
  var values = sheetValues_(sheet);
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

/**
 * One row by the id in column A (RID, PatientID, ...) without reading the sheet: a text finder on the column, then that
 * row's cells — two calls however long the sheet is. Same shape as a getRows_ entry (with _row); null if there is none.
 */
function getRowById_(sheet, headers, id) {
  var hit = sheet.getRange('A:A').createTextFinder(String(id)).matchEntireCell(true).matchCase(true).findNext();
  if (!hit || hit.getRow() < 2) return null;
  var r = hit.getRow(), vals = sheet.getRange(r, 1, 1, headers.length).getValues()[0], o = {};
  for (var c = 0; c < headers.length; c++) o[headers[c]] = vals[c];
  o._row = r;
  return o;
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
/** Column number -> letters (1 = A). */
function colLetter_(n) { var out = ''; while (n > 0) { var m = (n - 1) % 26; out = String.fromCharCode(65 + m) + out; n = Math.floor((n - 1) / 26); } return out; }
function a1_(row, col, numRows, numCols) { return colLetter_(col) + row + ':' + colLetter_(col + numCols - 1) + (row + numRows - 1); }

function applyTextColumnFormat_(sheet, headers, rowIndex, numRows) {
  var n = numRows || 1;
  var textRanges = (CONFIG.TEXT_COLUMNS[sheet.getName()] || []).map(function (col) { return headers.indexOf(col); })
    .filter(function (i) { return i > -1; }).map(function (i) { return a1_(rowIndex, i + 1, n, 1); });
  if (textRanges.length) sheet.getRangeList(textRanges).setNumberFormat('@'); // all text columns in ONE call
  resetNumberFormats_(sheet, headers, rowIndex, n);
}

/**
 * NUMBER_COLUMNS: resets any leftover (e.g. date) format on `numRows` rows
 * from `rowIndex` so numbers are stored as numbers — ONE call for all the
 * contiguous runs of numeric columns. Bulk writers (CSV import, seed) call
 * this directly for their whole block.
 */
function resetNumberFormats_(sheet, headers, rowIndex, numRows) {
  var numIdx = (CONFIG.NUMBER_COLUMNS[sheet.getName()] || []).map(function (c) { return headers.indexOf(c); })
    .filter(function (i) { return i > -1; }).sort(function (a, b) { return a - b; });
  var runs = [];
  for (var k = 0; k < numIdx.length; k++) {
    var start = numIdx[k];
    while (k + 1 < numIdx.length && numIdx[k + 1] === numIdx[k] + 1) k++;
    runs.push(a1_(rowIndex, start + 1, numRows, numIdx[k] - start + 1));
  }
  if (runs.length) sheet.getRangeList(runs).clearFormat();
}

function isDate_(v) { return Object.prototype.toString.call(v) === '[object Date]'; }

/** obj -> row values in header order; a NUMBER_COLUMNS value that was read back as a Date (see numOrNull_) goes back as its number. */
function rowValues_(sheet, headers, obj) {
  var numCols = CONFIG.NUMBER_COLUMNS[sheet.getName()] || [];
  return headers.map(function (h) {
    var v = obj[h] !== undefined ? obj[h] : '';
    return isDate_(v) && numCols.indexOf(h) > -1 ? numOrNull_(v) : v;
  });
}

/**
 * setValues that survives the sheet's row limit: a new sheet has 1000 rows, and writing below the last one fails, so
 * once the data reaches the bottom (StockLog gets there within weeks at a busy clinic) the sheet grows by 1000 rows.
 */
function setValuesGrowing_(sheet, rowIndex, numRows, numCols, values) {
  try {
    sheet.getRange(rowIndex, 1, numRows, numCols).setValues(values);
  } catch (e) {
    var missing = rowIndex + numRows - 1 - sheet.getMaxRows();
    if (missing <= 0) throw e;
    sheet.insertRowsAfter(sheet.getMaxRows(), Math.max(1000, missing));
    sheet.getRange(rowIndex, 1, numRows, numCols).setValues(values);
  }
}

function appendRow_(sheet, headers, obj) {
  var rowIndex = sheet.getLastRow() + 1, vals = rowValues_(sheet, headers, obj);
  applyTextColumnFormat_(sheet, headers, rowIndex);
  setValuesGrowing_(sheet, rowIndex, 1, headers.length, [vals]);
  cachePutRow_(sheet, rowIndex, vals);
  return rowIndex;
}

/** Appends several rows with one getLastRow, one format pass and one write (objs: an array of row objects). */
function appendRows_(sheet, headers, objs) {
  if (!objs.length) return [];
  var rowIndex = sheet.getLastRow() + 1, block = objs.map(function (o) { return rowValues_(sheet, headers, o); });
  applyTextColumnFormat_(sheet, headers, rowIndex, block.length);
  setValuesGrowing_(sheet, rowIndex, block.length, headers.length, block);
  block.forEach(function (vals, i) { cachePutRow_(sheet, rowIndex + i, vals); });
  return block.map(function (v, i) { return rowIndex + i; });
}

function updateRowByIndex_(sheet, headers, rowIndex, obj) {
  var vals = rowValues_(sheet, headers, obj);
  applyTextColumnFormat_(sheet, headers, rowIndex);
  sheet.getRange(rowIndex, 1, 1, headers.length).setValues([vals]);
  cachePutRow_(sheet, rowIndex, vals);
}

function deleteRowByIndex_(sheet, rowIndex) {
  sheet.deleteRow(rowIndex);
  if (_rowsCache_ && _rowsCache_[sheet.getName()]) _rowsCache_[sheet.getName()].splice(rowIndex - 1, 1);
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

/**
 * The Staff row for a StaffID. IDs aren't guaranteed unique in a live sheet
 * (a row added by hand can repeat one — S0001 for both admin and staff), so
 * the caller narrows it with the session's username or the sheet row number
 * listStaff handed out. Never falls back to "first match": an ambiguous or
 * stale lookup throws instead of touching someone else's account.
 */
function findStaffRow_(rows, staffId, by) {
  var matches = rows.filter(function (r) { return String(r.StaffID) === String(staffId); });
  if (by && by.username != null) matches = matches.filter(function (r) { return String(r.Username).toLowerCase() === String(by.username).toLowerCase(); });
  if (by && by.row != null && by.row !== '') matches = matches.filter(function (r) { return r._row === Number(by.row); });
  if (!matches.length) throw ValidationError_('ไม่พบผู้ใช้ (ข้อมูลอาจเปลี่ยนไปแล้ว กรุณาโหลดหน้าใหม่)');
  if (matches.length > 1) throw ValidationError_('รหัสพนักงาน ' + staffId + ' ซ้ำกันหลายแถว กรุณาโหลดหน้าใหม่แล้วลองอีกครั้ง');
  return matches[0];
}

function changePassword(token, currentPassword, newPassword) {
  return safeCall_(function () {
  var session = requireAuth_(token, 'staff');
  if (!newPassword) throw ValidationError_('กรุณากรอกรหัสผ่านใหม่');
  var sheet = getSheet_(CONFIG.SHEETS.STAFF);
  var user = findStaffRow_(getRows_(sheet), session.staffId, { username: session.username });
  var rowIndex = user._row;
  var currentHash = hashPassword_(currentPassword, user.PasswordSalt);
  if (currentHash !== user.PasswordHash) return { ok: false, message: 'รหัสผ่านเดิมไม่ถูกต้อง' };
  var salt = Utilities.getUuid();
  updateRowByIndex_(sheet, CONFIG.HEADERS.Staff, rowIndex, Object.assign({}, user, {
    PasswordHash: hashPassword_(newPassword, salt), PasswordSalt: salt
  }));
  return { ok: true };
  });
}

/** Admin-only: reset a staff member's password back to the default 12345. */
function resetStaffPassword(token, staffId, row) {
  return safeCall_(function () {
    var session = requireAuth_(token, 'admin');
    var sheet = getSheet_(CONFIG.SHEETS.STAFF);
    var user = findStaffRow_(getRows_(sheet), staffId, { row: row });
    var salt = Utilities.getUuid();
    updateRowByIndex_(sheet, CONFIG.HEADERS.Staff, user._row, Object.assign({}, user, {
      PasswordHash: hashPassword_('12345', salt), PasswordSalt: salt
    }));
    invalidateMasterCache_();
    writeAudit_(session, 'UPDATE', 'Staff', staffId, 'reset password: ' + user.Username);
    return { ok: true, message: 'รีเซ็ตรหัสผ่านของ ' + user.Username + ' เป็น 12345 เรียบร้อยแล้ว' };
  });
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
  var procRows = getRows_(getSheet_(CONFIG.SHEETS.PROCEDURES)).filter(function (p) { return isActiveStatus_(p.Status); });
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
    staff: staff.map(function (s) { return { id: s.StaffID, name: s.Name, username: s.Username, level: s.Level, active: !isFalse_(s.Active) }; }),
    treatmentCategories: treatmentCats.map(function (c) { return { id: c.CategoryID, name: c.Name }; }),
    stockCategories: stockCats.map(function (c) { return { id: c.CategoryID, name: c.Name }; }),
    expenseCategories: expenseCats.map(function (c) { return { id: c.CategoryID, name: c.Name }; }),
    promotions: promoRows.map(function (p) { return { id: p.PromoID, name: p.Name, price: Number(p.Price) || 0, doctorSharePct: Number(p.DoctorSharePct) || 0 }; }),
    procedures: procRows.map(function (p) { return { id: p.ProcID, name: p.Name, category: p.Category || '', price: Number(p.Price) || 0 }; }),
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
    var data = rows.map(function (p) { return { id: p.ID, hn: p.HN || '', firstName: p.FirstName, lastName: p.LastName, nickname: p.Nickname || '', phone: p.Phone }; });
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
      return { id: s.ItemID, name: s.Name, category: s.Category, unit: s.Unit, qty: Number(s.Qty) || 0, type: stockTypeOf_(s), prices: STOCK_PRICE_TIERS.map(function (t) { return numOrNull_(s['Price' + t]); }) };
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
    if (filters.q) rows = rows.filter(patientMatcher_(filters.q));
    rows.sort(function (a, b) { return String(b.CreatedAt).localeCompare(String(a.CreatedAt)); });
    var page = paginate_(rows, opts.page, opts.pageSize);
    page.items = page.items.map(mapPatientOut_);
    return { ok: true, data: page };
  });
}

/** Patients-row filter for a search box: name, phone digits or HN. */
function patientMatcher_(q) {
  q = String(q || '').trim().toLowerCase();
  var qDigits = q.replace(/[^\d]/g, '');
  return function (p) {
    var fullName = (p.FirstName + ' ' + p.LastName).toLowerCase();
    return fullName.indexOf(q) > -1 || (qDigits && String(p.Phone).indexOf(qDigits) > -1) ||
      (p.HN && String(p.HN).toLowerCase().indexOf(q) > -1);
  };
}

function mapPatientOut_(p) {
  return {
    id: p.ID, hn: p.HN || '', firstName: p.FirstName, lastName: p.LastName, nickname: p.Nickname || '', phone: p.Phone, dob: dateStr_(p.DOB),
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
      return s + (Number(r.Amount) || 0) + (Number(r.PromoAmount) || 0) + (useCanula ? (Number(r.CanulaAmount) || 0) : 0) + (Number(r.MedicineAmount) || 0) + (Number(r.OralMedAmount) || 0) + (Number(r.OtherAmount) || 0);
    }, 0);
    return {
      ok: true,
      data: {
        patient: mapPatientOut_(p),
        visitCount: visitCount,
        totalSpent: totalSpent,
        records: records.map(mapRecordOut_), // the panel pages these 15 at a time
        deposits: patientDeposits_(id)
      }
    };
  });
}

/** Lightweight suggest list for search-as-you-type widgets (used by Records/Receipt patient search too). */
function searchPatients(token, q, limit) {
  return safeCall_(function () {
    requireAuth_(token, 'staff');
    if (!String(q || '').trim()) return { ok: true, data: [] };
    var rows = getRows_(getSheet_(CONFIG.SHEETS.PATIENTS));
    var matches = rows.filter(patientMatcher_(q)).slice(0, limit || 8);
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

/** A patient made on the spot while dispensing (name only, no phone yet): the typed name followed by this tag, until someone edits the history. */
var PENDING_NAME_TAG_ = '(รออัพเดท)';
function patientFullName_(p) { return (String(p.FirstName || '') + ' ' + String(p.LastName || '')).trim(); }
function isPendingPatientName_(name) { return String(name || '').indexOf(PENDING_NAME_TAG_) > -1; }
/** Adds the placeholder patient for a typed name (inside the caller's lock); returns its Patients row object. */
function createPendingPatient_(session, typedName) {
  var name = String(typedName || '').replace(/\s+/g, ' ').trim();
  if (!name) throw ValidationError_('กรุณาเลือกคนไข้หรือพิมพ์ชื่อคนไข้');
  if (name.length > 120) throw ValidationError_('ชื่อคนไข้ยาวเกินไป');
  var obj = {
    ID: 'P' + nextSeq_('PATIENT', 5), HN: '', FirstName: isPendingPatientName_(name) ? name : name + ' ' + PENDING_NAME_TAG_, LastName: '', Nickname: '', Phone: '',
    DOB: '', Gender: '', Email: '', PhotoFileId: '', JoinedDate: todayStr_(), CreatedAt: nowStr_(), UpdatedAt: nowStr_()
  };
  appendRow_(getSheet_(CONFIG.SHEETS.PATIENTS), CONFIG.HEADERS.Patients, obj);
  writeAudit_(session, 'CREATE', 'Patient', obj.ID, obj.FirstName + ' (จากจัดยา-เบิกยา)');
  invalidatePatientsLiteCache_();
  return Object.assign({ _row: 0 }, obj);
}
/** Carries a patient's new name/phone onto the rows that copied them (Records, MedWithdrawals, MedDeposits). */
function renamePatientRefs_(patientId, name, phone) {
  [[CONFIG.SHEETS.RECORDS, CONFIG.HEADERS.Records, true], [CONFIG.SHEETS.MED_WITHDRAWALS, CONFIG.HEADERS.MedWithdrawals, true], [CONFIG.SHEETS.MED_DEPOSITS, CONFIG.HEADERS.MedDeposits, false]].forEach(function (t) {
    var sheet = getSheet_(t[0]);
    getRows_(sheet).filter(function (r) { return r.PatientID === patientId; }).forEach(function (r) {
      var next = Object.assign({}, r, { PatientName: name });
      if (t[2]) next.Phone = phone;
      updateRowByIndex_(sheet, t[1], r._row, next);
    });
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
      var hn = String(patient.hn || '').trim();
      if (hn) {
        var taken = getRows_(sheet).filter(function (p) { return String(p.HN || '').trim().toLowerCase() === hn.toLowerCase() && p.ID !== patient.id; })[0];
        if (taken) throw ValidationError_('เลข HN ' + hn + ' ซ้ำกับคนไข้ ' + taken.FirstName + ' ' + taken.LastName);
      }
      var obj = {
        ID: isNew ? ('P' + nextSeq_('PATIENT', 5)) : patient.id, HN: hn,
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
        var before = getRows_(sheet).filter(function (p) { return p.ID === patient.id; })[0];
        updateRowByIndex_(sheet, CONFIG.HEADERS.Patients, rowIndex, obj);
        // A patient made while dispensing (รออัพเดท) has its name on records / withdrawals / deposits already: bring them up to date.
        if (before && isPendingPatientName_(patientFullName_(before))) renamePatientRefs_(obj.ID, patientFullName_(obj), obj.Phone);
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

/* ---------------------------------------------------------------- */
/* รายการหัตถการ (จัดการสินค้า/บริการ) — the record modal's หัตถการ   */
/* picker reads the active ones from master data, so every write here   */
/* invalidates it, same as Promotions above.                             */
/* ---------------------------------------------------------------- */
function mapProcedureOut_(p) {
  return { id: p.ProcID, name: p.Name, category: p.Category || '', price: Number(p.Price) || 0, status: p.Status || 'active' };
}

function listProcedures(token) {
  return safeCall_(function () {
    requireAuth_(token, 'staff');
    var rows = getRows_(getSheet_(CONFIG.SHEETS.PROCEDURES));
    rows.sort(function (a, b) { return String(a.Category).localeCompare(String(b.Category), 'th') || String(a.Name).localeCompare(String(b.Name), 'th'); });
    return { ok: true, data: rows.map(mapProcedureOut_) };
  });
}

function saveProcedure(token, proc) {
  return safeCall_(function () {
    var session = requireAuth_(token, 'admin');
    requireFields_(proc, { name: 'ชื่อหัตถการ' });
    var sheet = getSheet_(CONFIG.SHEETS.PROCEDURES);
    var isNew = !proc.id;
    var existing = isNew ? null : getRows_(sheet).filter(function (r) { return r.ProcID === proc.id; })[0];
    if (!isNew && !existing) throw ValidationError_('ไม่พบข้อมูลหัตถการ');
    var obj = {
      ProcID: isNew ? ('PR' + nextSeq_('PROCEDURE', 4)) : proc.id,
      Name: String(proc.name).trim(), Category: proc.category || '',
      Price: requireNumber_(proc.price || 0, 'ราคา', { min: 0 }),
      Status: proc.status === 'inactive' ? 'inactive' : 'active',
      CreatedAt: isNew ? nowStr_() : existing.CreatedAt, UpdatedAt: nowStr_()
    };
    if (isNew) appendRow_(sheet, CONFIG.HEADERS.Procedures, obj);
    else updateRowByIndex_(sheet, CONFIG.HEADERS.Procedures, existing._row, obj);
    invalidateMasterCache_();
    writeAudit_(session, isNew ? 'CREATE' : 'UPDATE', 'Procedure', obj.ProcID, obj.Name);
    return { ok: true, data: mapProcedureOut_(obj) };
  });
}

function toggleProcedureStatus(token, id) {
  return safeCall_(function () {
    var session = requireAuth_(token, 'admin');
    var sheet = getSheet_(CONFIG.SHEETS.PROCEDURES);
    var row = getRows_(sheet).filter(function (r) { return r.ProcID === id; })[0];
    if (!row) throw ValidationError_('ไม่พบข้อมูลหัตถการ');
    var newStatus = row.Status === 'inactive' ? 'active' : 'inactive';
    updateRowByIndex_(sheet, CONFIG.HEADERS.Procedures, row._row, Object.assign({}, row, { Status: newStatus, UpdatedAt: nowStr_() }));
    invalidateMasterCache_();
    writeAudit_(session, 'UPDATE', 'Procedure', id, 'status -> ' + newStatus);
    return { ok: true, data: { status: newStatus } };
  });
}

function deleteProcedure(token, id) {
  return safeCall_(function () {
    var session = requireAuth_(token, 'admin');
    var sheet = getSheet_(CONFIG.SHEETS.PROCEDURES);
    var rowIndex = findRowIndexById_(sheet, 'ProcID', id);
    if (rowIndex < 0) throw ValidationError_('ไม่พบข้อมูลหัตถการ');
    deleteRowByIndex_(sheet, rowIndex);
    invalidateMasterCache_();
    writeAudit_(session, 'DELETE', 'Procedure', id, '');
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
      return { id: s.StaffID, name: s.Name, role: s.Role, username: s.Username, lineId: s.LineID, level: s.Level, active: !isFalse_(s.Active), row: s._row };
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
      // Skip IDs already in the sheet, in case the STAFF counter is behind hand-added rows.
      var staffId;
      do { staffId = 'S' + nextSeq_('STAFF', 4); } while (rows.some(function (s) { return String(s.StaffID) === staffId; }));
      var obj = {
        StaffID: staffId, Name: staff.name, Role: staff.role || '', Username: staff.username,
        PasswordHash: hashPassword_('12345', salt), PasswordSalt: salt, LineID: staff.lineId || '',
        Level: staff.level, Active: 'TRUE', CreatedAt: nowStr_()
      };
      appendRow_(sheet, CONFIG.HEADERS.Staff, obj);
      invalidateMasterCache_();
      writeAudit_(session, 'CREATE', 'Staff', obj.StaffID, obj.Name);
      return { ok: true, data: obj, message: 'สร้างผู้ใช้ใหม่แล้ว รหัสผ่านเริ่มต้นคือ 12345' };
    } else {
      var existing = findStaffRow_(rows, staff.id, { row: staff.row });
      var rowIndex = existing._row;
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

function deleteStaff(token, id, row) {
  return safeCall_(function () {
    var session = requireAuth_(token, 'admin');
    var sheet = getSheet_(CONFIG.SHEETS.STAFF);
    var user = findStaffRow_(getRows_(sheet), id, { row: row });
    if (String(user.Username).toLowerCase() === String(session.username).toLowerCase()) throw ValidationError_('ไม่สามารถลบบัญชีของตัวเองได้');
    deleteRowByIndex_(sheet, user._row);
    invalidateMasterCache_();
    writeAudit_(session, 'DELETE', 'Staff', id, user.Username);
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
var STOCK_LOG_KINDS = { update: 1, in: 1, out: 1, set: 1, import: 1, opening: 1, record: 1, withdraw: 1, purchase: 1 };

function listStock(token, opts) {
  return safeCall_(function () {
    var session = requireAuth_(token, 'staff');
    opts = opts || {};
    var filters = compactFilters_(opts.filters);
    var rows = getRows_(getSheet_(CONFIG.SHEETS.STOCK));
    if (filters.type) rows = rows.filter(function (r) { return stockTypeOf_(r) === filters.type; });
    // Stat cards count the whole tab (before search/category filters). ใกล้หมด
    // uses the item's MinQty when set, else the old fixed threshold of 10 —
    // same rule as the row badge on the client.
    var summary = { total: rows.length, low: 0, out: 0 };
    rows.forEach(function (r) {
      var qty = Number(r.Qty) || 0, min = numOrNull_(r.MinQty);
      if (qty <= 0) summary.out++;
      else if (min != null ? qty <= min : qty <= 10) summary.low++;
    });
    if (filters.q) {
      var q = String(filters.q).trim().toLowerCase();
      rows = rows.filter(function (r) { return String(r.Name).toLowerCase().indexOf(q) > -1; });
    }
    if (filters.category) rows = rows.filter(function (r) { return r.Category === filters.category; });

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
    // จำนวนฝาก = ยาฝาก still held for patients (MedDeposits). Dispensing already
    // took those units out of Stock.Qty, so Qty is what's พร้อมใช้ and the two
    // together are what's physically at the clinic.
    var depositByItem = {};
    if (filters.type !== 'supply' && filters.type !== 'oral') {
      getRows_(getSheet_(CONFIG.SHEETS.MED_DEPOSITS)).forEach(function (d) {
        var q = Number(d.QtyRemaining) || 0;
        if (q > 0) depositByItem[d.ItemID] = round2_((depositByItem[d.ItemID] || 0) + q);
      });
    }
    page.items = page.items.map(function (r) {
      var o = mapStockOut_(r, latestByItem[r.ItemID], showCost);
      o.depositQty = depositByItem[r.ItemID] || 0;
      return o;
    });
    page.summary = summary;
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

/** 'medicine' (ยา), 'oral' (ยาทาน) or 'supply' (เวชภัณฑ์) — rows saved before StockType existed are blank and count as 'supply'. */
function stockTypeOf_(s) {
  return stockTypeFrom_(s.StockType);
}
function stockTypeFrom_(t) {
  return t === 'medicine' ? 'medicine' : (t === 'oral' ? 'oral' : 'supply');
}

/** Blank cell -> null (so the client can show "-"), anything else -> Number. */
function numOrNull_(v) {
  if (v === '' || v === null || v === undefined) return null;
  // A number written into a date-formatted cell is stored as a date serial
  // (10 -> 1900-01-09) and read back as a Date — recover the number (days
  // since 1899-12-30), snapping tiny old-timezone offsets to whole numbers.
  if (isDate_(v)) {
    var days = (v.getTime() - new Date(1899, 11, 30).getTime()) / 864e5;
    return Math.abs(days - Math.round(days)) < 0.05 ? Math.round(days) : round2_(days);
  }
  return Number(v) || 0;
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
        StockType: stockTypeFrom_(item.type),
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
      else if (payload.kind === 'out') { qtyOut = qty; remaining = current - qty; } // may go negative, same as treatment records
      else {
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
    var type = stockTypeFrom_(payload.type);
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
      // Numeric cells that were read back as dates (see numOrNull_) are written back as numbers.
      var numCols = CONFIG.NUMBER_COLUMNS.Stock;
      var toValues = function (obj) { return headers.map(function (h) { var v = obj[h] !== undefined ? obj[h] : ''; return isDate_(v) && numCols.indexOf(h) > -1 ? numOrNull_(v) : v; }); };
      if (allRows.length) {
        var lastRow = stockSheet.getLastRow();
        var block = stockSheet.getRange(2, 1, lastRow - 1, headers.length).getValues();
        allRows.forEach(function (r) { block[r._row - 2] = toValues(r); });
        resetNumberFormats_(stockSheet, headers, 2, lastRow - 1);
        stockSheet.getRange(2, 1, lastRow - 1, headers.length).setValues(block);
      }
      if (created.length) {
        var firstNew = stockSheet.getLastRow() + 1;
        resetNumberFormats_(stockSheet, headers, firstNew, created.length);
        stockSheet.getRange(firstNew, 1, created.length, headers.length).setValues(created.map(toValues));
      }
      if (logs.length) {
        var logIds = nextSeqBlock_('STOCKLOG', logs.length, 5);
        var logSheet = getSheet_(CONFIG.SHEETS.STOCK_LOG);
        var logValues = logs.map(function (l, i) {
          var obj = stockLogRow_(session, 'SL' + logIds[i], today, l.row, 'import', l.qtyIn, l.qtyOut, l.remaining, 'นำเข้า CSV');
          return CONFIG.HEADERS.StockLog.map(function (h) { return obj[h] !== undefined ? obj[h] : ''; });
        });
        var firstLog = logSheet.getLastRow() + 1;
        resetNumberFormats_(logSheet, CONFIG.HEADERS.StockLog, firstLog, logValues.length);
        logSheet.getRange(firstLog, 1, logValues.length, CONFIG.HEADERS.StockLog.length).setValues(logValues);
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
    requireNumber_(entry.qtyRemaining, 'จำนวนคงเหลือ'); // may be negative (stock is allowed below zero)

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
   MedicineJSON: "รายการยา" — JSON array of {withdrawId?, itemId, name,
     qtyDispensed, qtyUsed, qtyDeposit, expiry, unitPrice, amount,
     unitCost?}: a line with withdrawId charges that จัดยา-เบิกยา line (its
     qty already left Stock); any other line is เบิกเพิ่ม and its
     qtyDispensed is taken out of Stock here. qtyDeposit (= dispensed -
     used) becomes a ยาฝาก lot expiring on `expiry`, amount = qtyDispensed *
     unitPrice is charged. Lines saved before this shape existed are
     {name, amount}.
   OwnMedJSON {name,qty} "รายการถือยาเข้า", DepositUseJSON {itemId,name,qty,
     lots} "รายการใช้ยาฝาก" (taken out of MedDeposits lots, see
     planDeposits_), MessengerJSON {name,qty,note} "รายการ Messenger/พัสดุ"
     — none of these three are charged.
   ProcedureJSON: "รายการหัตถการ" — {procId,name,price,amount}; Amount is
     their sum.
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

function stripCost_(l) { var o = Object.assign({}, l); delete o.unitCost; return o; }
/** The record's Canula lines: CanulaJSON, or (rows from before it) the one CanulaAmount as a single line. */
function canulaItemsOf_(r) {
  var list = parseLineItems_(r.CanulaJSON);
  if (list.length || !isTrue_(r.UseCanula)) return list;
  return Number(r.CanulaAmount) > 0 ? [{ name: '', qty: 1, amount: Number(r.CanulaAmount) }] : [];
}

function mapRecordOut_(r) {
  var useCanula = String(r.UseCanula) === 'TRUE' || r.UseCanula === true;
  var canulaAmount = Number(r.CanulaAmount) || 0;
  var amount = Number(r.Amount) || 0;
  var otherAmount = Number(r.OtherAmount) || 0;
  var promoAmount = Number(r.PromoAmount) || 0;
  var medicineAmount = Number(r.MedicineAmount) || 0;
  var doctorShareRaw = Number(r.DoctorShareRaw) || 0;
  var salesShare = Number(r.SalesShare) || 0;
  // grossBase = procedure + promo + canula (what the sales/doctor % apply
  // to); collected adds ยาเบิกเพิ่ม + รายการอื่นๆ, and ส่วนแบ่งคลีนิค is all of
  // it minus the sales and doctor shares (see recordClinicShare_).
  var grossBase = round2_(amount + promoAmount + (useCanula ? canulaAmount : 0));
  var collected = recordCollected_(r);
  var medicineSplit = splitMedicineLines_(r);
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
    // Revenue-split fields (see buildDoctorShares_/recomputeDoctorDayShares_).
    // doctorShareRaw/clinicShareRaw are the plain per-record commission split
    // BEFORE the doctor's daily-minimum-guarantee top-up is applied — the
    // per-record detail panel shows these (not doctorShare/clinicShare, which
    // are the topped-up, actually-paid-out amounts) since the guarantee only
    // makes sense aggregated across a whole day; see computeDailySummaryDetail_.
    otherAmount: otherAmount,
    salesCommissionPct: r.SalesCommissionPct === '' || r.SalesCommissionPct == null ? null : Number(r.SalesCommissionPct) || 0,
    salesShare: salesShare,
    doctorCommissionPct: Number(r.DoctorCommissionPct) || 0,
    doctorShare: Number(r.DoctorShare) || 0,
    doctorShareRaw: doctorShareRaw,
    doctorShareNote: r.DoctorShareNote || '',
    doctorShares: recordDoctorShares_(r),
    clinicShare: recordClinicShare_(r),
    clinicShareRaw: round2_(collected - salesShare - doctorShareRaw),
    // Round 13 line-item groups.
    usePromo: String(r.UsePromo) === 'TRUE' || r.UsePromo === true,
    promoItems: parseLineItems_(r.PromoJSON), promoAmount: promoAmount,
    // The two ยา groups of the form (see splitMedicineLines_): ไม่คิดค่ายา = medicineItems, คิดค่ายา = chargedMedItems.
    // unitCost (ราคาทุน) stays on the server — only รายงานรายเซลล์ reads it.
    useMedicine: medicineSplit.uncharged.length > 0, medicineItems: medicineSplit.uncharged.map(stripCost_),
    useChargedMed: medicineSplit.charged.length > 0, chargedMedItems: medicineSplit.charged.map(stripCost_),
    medicineAmount: medicineAmount,
    canulaItems: canulaItemsOf_(r),
    useOralMed: isTrue_(r.UseOralMed), oralMedItems: parseLineItems_(r.OralMedJSON), oralMedAmount: Number(r.OralMedAmount) || 0,
    useProcedure: isTrue_(r.UseProcedure), procedureItems: parseLineItems_(r.ProcedureJSON),
    useOther: String(r.UseOther) === 'TRUE' || r.UseOther === true,
    otherItems: parseLineItems_(r.OtherItemsJSON),
    useOwnMed: isTrue_(r.UseOwnMed), ownMedItems: parseLineItems_(r.OwnMedJSON),
    useDepositUse: isTrue_(r.UseDepositUse), depositUseItems: parseLineItems_(r.DepositUseJSON),
    useMessenger: isTrue_(r.UseMessenger), messengerItems: parseLineItems_(r.MessengerJSON),
    stockApplied: isTrue_(r.StockApplied),
    withdrawItems: parseLineItems_(r.WithdrawJSON),
    estimatedRevenue: grossBase,
    totalCollected: collected,
    visitStatus: visitStatusOf_(r), // see visitStatusOf_: older ''/'scheduled' read as จัดยาแล้ว
    queueDate: r.QueueDate ? dateStr_(r.QueueDate) : '', queueDoctor: Number(r.QueueDoctor) || null, queuePay: null,
    visitDoctorId: r.VisitDoctorID || '', visitDoctorName: r.VisitDoctorName || '',
    // Round 14 item 1.7 — OPD photo attachment (Drive fileId; fetch its
    // viewable data URI via getPhotoUrl, same as patient/login/logo photos).
    opdImageFileId: r.OpdImageFileId || ''
  };
}

/* ======================================================================
   REVENUE SPLIT ENGINE — see the spec's formulas (ยอดชำระจากลูกค้า /
   ส่วนแบ่งเซลล์ / ส่วนแบ่งแพทย์ / ส่วนแบ่งคลินิก) and the confirmed
   clarifications: (1) "สถานะ" (payment: ชำระแล้ว/รอชำระ/ยกเลิก) stays as-is
   and still gates revenue inclusion; VisitStatus (จัดยาแล้ว/เสร็จสิ้น) is a separate, independent workflow marker. (2)
   OtherAmount/MedicineAmount ("จำนวนเงินอื่นๆ"/"ค่ายาเบิกเพิ่ม") are not in the
   sales/doctor % base; they now go to the clinic (there is no supplier share).
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
     collected       = A + P + K + M + O                        -- M ยาเบิกเพิ่ม, O รายการอื่นๆ
     clinicShareRaw  = collected - salesShare - doctorShareRaw  -- (3) everything left over
   Since multi-doctor records, A and P are per line: each หัตถการ line pays
   its own doctor A_line * Cd_line, each โปร line pays its doctor
   P_line * promo%, and the daily minimum is per doctor (buildDoctorShares_ /
   recomputeDoctorDayShares_).
   ====================================================================== */

/** Checkbox-style flags are written as 'TRUE'/'FALSE', but Sheets stores them as real booleans — accept both. */
function isTrue_(v) {
  return v === true || String(v).toUpperCase() === 'TRUE';
}

/** Explicitly FALSE (a blank cell is not): Staff.Active treats blank as active. */
function isFalse_(v) {
  return v === false || String(v).toUpperCase() === 'FALSE';
}

function round2_(n) {
  return Math.round((Number(n) || 0) * 100) / 100;
}

/** What the patient pays for one Records row: หัตถการ + โปร + canula + ซื้อยาเพิ่ม (คิดค่ายา) + จ่ายยาทาน + รายการอื่นๆ. */
function recordCollected_(r) {
  var canula = isTrue_(r.UseCanula) ? (Number(r.CanulaAmount) || 0) : 0;
  return round2_((Number(r.Amount) || 0) + (Number(r.PromoAmount) || 0) + canula + (Number(r.MedicineAmount) || 0) + (Number(r.OralMedAmount) || 0) + (Number(r.OtherAmount) || 0));
}

/* ---------------------------------------------------------------------
   สถานะการรักษา: จัดยาแล้ว → รอพบแพทย์ → พบแพทย์ → รอชำระเงิน → เสร็จสิ้น.
   Front moves จัดยาแล้ว → รอพบแพทย์ → พบแพทย์ (choosing the doctor) and takes payment;
   the doctor moves พบแพทย์ → รอชำระเงิน by saving the treatment record. Back
   while unpaid; เสร็จสิ้น is only ever reached by payment. Arriving (รอพบแพทย์)
   takes the day's next queue number, which the customer keeps to the end.
   --------------------------------------------------------------------- */
var VISIT_FLOW_ = ['dispensed', 'waiting_doctor', 'with_doctor', 'waiting_payment', 'completed'];

/** The record's step, whatever older rows stored: paid = เสร็จสิ้น; a stored 'completed' that isn't paid is รอชำระเงิน; ''/'scheduled' = จัดยาแล้ว. */
function visitStatusOf_(r) {
  if (r.Status === 'paid') return 'completed';
  var v = String(r.VisitStatus || '');
  if (v === 'waiting_doctor' || v === 'with_doctor' || v === 'waiting_payment') return v;
  if (v === 'completed') return r.Status === 'cancelled' ? 'completed' : 'waiting_payment';
  return 'dispensed';
}

/** The next queue number of `date` for kind 'D' (พบแพทย์) or 'P' (ชำระเงิน); starts at 1 each day. */
function nextQueueNo_(kind, date) { return Number(nextSeq_('QUEUE_' + kind + '_' + date, 1)); }

/**
 * Moves a Records row object (not yet written) to step `to`, assigning the queue number of `day` on arrival.
 * doctor ({id, name}): who the customer sees, required when going to พบแพทย์ unless the record already has one.
 * shiftDate: a record dispensed on an earlier day takes `day` as its treatment date when the customer arrives.
 * Refused once paid, and เสร็จสิ้น itself is never a target (payment does that).
 */
function applyVisitStatus_(r, to, day, shiftDate, doctor) {
  var from = visitStatusOf_(r);
  if (from === to) return;
  if (VISIT_FLOW_.indexOf(to) < 0) throw ValidationError_('สถานะการรักษาไม่ถูกต้อง');
  if (r.Status === 'paid') throw ValidationError_('ชำระเงินแล้ว — เปลี่ยนสถานะไม่ได้ (ยกเลิกการชำระเงินก่อน)');
  if (to === 'completed') throw ValidationError_('สถานะเสร็จสิ้นเกิดจากการชำระเงินเท่านั้น');
  if (to === 'with_doctor') {
    if (doctor) { r.VisitDoctorID = doctor.id; r.VisitDoctorName = doctor.name; }
    else if (!r.VisitDoctorID) throw ValidationError_('กรุณาเลือกแพทย์ที่จะเข้าพบ');
  } else if (to === 'dispensed' || to === 'waiting_doctor') {
    r.VisitDoctorID = ''; r.VisitDoctorName = '';
  }
  if (to === 'dispensed') {
    r.QueueDate = ''; r.QueueDoctor = ''; r.QueuePay = '';
  } else {
    // First arrival: the customer takes today as their queue day (and treatment date, when asked) and a doctor-queue number.
    if (!r.QueueDate) {
      if (shiftDate && dateStr_(r.Date) !== day) r.Date = day;
      r.QueueDate = day;
    }
    if (!r.QueueDoctor) r.QueueDoctor = nextQueueNo_('D', dateStr_(r.QueueDate));
    r.QueuePay = '';
  }
  r.VisitStatus = to;
}

/** ส่วนแบ่งคลีนิค = everything collected − ส่วนแบ่งเซลล์ − ส่วนแบ่งแพทย์ (final, after the daily minimum). Read this rather than the ClinicShare column, which older rows wrote without ยา/รายการอื่นๆ. */
function recordClinicShare_(r) {
  return round2_(recordCollected_(r) - (Number(r.SalesShare) || 0) - (Number(r.DoctorShare) || 0));
}

/** The record's per-doctor shares ({doctorId, name, pct, raw, share, note}); rows from before DoctorSharesJSON have their one doctor in the plain columns. */
function recordDoctorShares_(r) {
  var list = parseLineItems_(r.DoctorSharesJSON);
  if (list.length || r.DoctorSharesJSON) return list;
  return r.DoctorID ? [{
    doctorId: r.DoctorID, name: r.DoctorName || '', pct: Number(r.DoctorCommissionPct) || 0,
    raw: Number(r.DoctorShareRaw) || 0, share: Number(r.DoctorShare) || 0, note: r.DoctorShareNote || ''
  }] : [];
}

/** Writes a per-doctor share list onto a Records row object (in place) and derives the summary columns from it. */
function applyDoctorShares_(r, list) {
  r.DoctorSharesJSON = JSON.stringify(list);
  r.DoctorID = list.length ? list[0].doctorId : '';
  r.DoctorName = list.map(function (d) { return d.name; }).join(', ');
  r.DoctorCommissionPct = list.length ? list[0].pct : '';
  r.DoctorShareRaw = round2_(list.reduce(function (s, d) { return s + (Number(d.raw) || 0); }, 0));
  r.DoctorShare = round2_(list.reduce(function (s, d) { return s + (Number(d.share) || 0); }, 0));
  r.DoctorShareNote = list.map(function (d) { return d.note; }).filter(Boolean).join(' · ');
  r.ClinicShare = recordClinicShare_(r);
  return r;
}

/**
 * Raw (pre-daily-minimum) ส่วนแบ่งแพทย์ per doctor. entries: one per
 * หัตถการ line {doctorId, base: amount} (the doctor's own CommissionPct
 * applies) and one per โปร line {doctorId, extra: amount * promo
 * DoctorSharePct / 100}. Returns [{doctorId, name, pct, raw, share, note}]
 * in first-appearance order; share = raw until recomputeDoctorDayShares_
 * applies the minimum.
 */
function buildDoctorShares_(entries, doctorById) {
  var byId = {}, order = [];
  entries.forEach(function (e) {
    if (!e.doctorId) return;
    var g = byId[e.doctorId];
    if (!g) {
      var d = doctorById[e.doctorId];
      g = byId[e.doctorId] = { doctorId: e.doctorId, name: d ? d.Name : '', pct: d ? Number(d.CommissionPct) || 0 : 0, raw: 0, share: 0, note: '' };
      order.push(g);
    }
    g.raw += (Number(e.base) || 0) * g.pct / 100 + (Number(e.extra) || 0);
  });
  order.forEach(function (g) { g.raw = round2_(g.raw); g.share = g.raw; });
  return order;
}

/**
 * Enforces one doctor's daily minimum guarantee across every (non-cancelled)
 * record on one date that pays that doctor: if the doctor's raw shares for
 * the day add up to less than their MinAmount, the shortfall is topped up
 * onto their entry in the most recently saved record (with a note), and
 * their entries in the day's other records go back to plain raw. The
 * record's totals and ClinicShare are re-derived (applyDoctorShares_).
 * Must run inside the caller's script lock.
 *
 * `allRows` is an optional Records getRows_() snapshot taken after the
 * caller's own write; rows this rewrites are updated in it in place, so the
 * caller can run this for several doctors on one snapshot (see saveRecord).
 * Returns the day's group for that doctor.
 */
function recomputeDoctorDayShares_(doctorId, date, allRows) {
  if (!doctorId || !date) return [];
  var doctorRows = getRows_(getSheet_(CONFIG.SHEETS.DOCTORS));
  var doctor = doctorRows.filter(function (d) { return d.DoctorID === doctorId; })[0];
  var minAmount = doctor ? (Number(doctor.MinAmount) || 0) : 0;

  var sheet = getSheet_(CONFIG.SHEETS.RECORDS);
  var rows = allRows || getRows_(sheet);
  var entryOf = function (list) { return list.filter(function (d) { return d.doctorId === doctorId; })[0]; };
  var group = rows.filter(function (r) {
    return dateStr_(r.Date) === date && r.Status !== 'cancelled' && entryOf(recordDoctorShares_(r));
  });
  if (!group.length) return [];
  group.sort(function (a, b) { return String(a.CreatedAt).localeCompare(String(b.CreatedAt)); });

  var rawTotal = 0;
  group.forEach(function (r) { rawTotal += Number(entryOf(recordDoctorShares_(r)).raw) || 0; });
  var shortfall = minAmount > rawTotal ? round2_(minAmount - rawTotal) : 0;
  var lastIdx = group.length - 1;

  group.forEach(function (r, i) {
    var list = recordDoctorShares_(r).map(function (d) { return Object.assign({}, d); });
    var e = entryOf(list);
    e.share = Number(e.raw) || 0;
    e.note = '';
    if (shortfall > 0 && i === lastIdx) {
      e.share = round2_(e.share + shortfall);
      e.note = 'ปรับเพิ่มให้ครบยอดขั้นต่ำต่อวัน (' + minAmount.toLocaleString('en-US') + ' บาท) ของ ' + (doctor ? doctor.Name : doctorId);
    }
    var before = [r.DoctorSharesJSON, Number(r.DoctorShare), String(r.DoctorShareNote || ''), Number(r.ClinicShare)].join('|');
    applyDoctorShares_(r, list);
    if ([r.DoctorSharesJSON, Number(r.DoctorShare), String(r.DoctorShareNote || ''), Number(r.ClinicShare)].join('|') !== before) {
      updateRowByIndex_(sheet, CONFIG.HEADERS.Records, r._row, r);
    }
  });
  return group;
}

/** Re-applies the daily minimum for every doctor a record pays (before and after a change) on its date(s). */
function recomputeRecordDoctors_(oldRow, newRow, allRows) {
  var seen = {};
  [oldRow, newRow].forEach(function (r) {
    if (!r) return;
    var date = dateStr_(r.Date);
    recordDoctorShares_(r).forEach(function (d) {
      var key = d.doctorId + '|' + date;
      if (seen[key]) return;
      seen[key] = true;
      recomputeDoctorDayShares_(d.doctorId, date, allRows);
    });
  });
}

function listRecords(token, opts) {
  return safeCall_(function () {
    requireAuth_(token, 'staff');
    opts = opts || {};
    var filters = compactFilters_(opts.filters);
    var rows = getRows_(getSheet_(CONFIG.SHEETS.RECORDS));
    var hnById = patientHnById_();
    if (filters.q) {
      var q = String(filters.q).trim().toLowerCase();
      var qDigits = q.replace(/[^\d]/g, '');
      rows = rows.filter(function (r) {
        return String(r.PatientName).toLowerCase().indexOf(q) > -1 || (qDigits && String(r.Phone).indexOf(qDigits) > -1) ||
          String(hnById[r.PatientID] || '').toLowerCase().indexOf(q) > -1;
      });
    }
    if (filters.doctorId) rows = rows.filter(function (r) { return recordDoctorShares_(r).some(function (d) { return d.doctorId === filters.doctorId; }); });
    if (filters.status) rows = rows.filter(function (r) { return r.Status === filters.status; });
    if (filters.visitStatus) rows = rows.filter(function (r) { return visitStatusOf_(r) === filters.visitStatus; });
    if (filters.excludeVisitStatus) rows = rows.filter(function (r) { return visitStatusOf_(r) !== filters.excludeVisitStatus; });
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
    page.items = page.items.map(function (r) { var o = mapRecordOut_(r); o.hn = hnById[r.PatientID] || ''; return o; });
    return { ok: true, data: page };
  });
}

/** {PatientID: HN} for list screens (HN column and search). */
function patientHnById_() {
  var map = {};
  getRows_(getSheet_(CONFIG.SHEETS.PATIENTS)).forEach(function (p) { map[p.ID] = p.HN || ''; });
  return map;
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
    var useCanula = isTrue_(r.UseCanula);
    var canula = useCanula ? (Number(r.CanulaAmount) || 0) : 0;
    var other = Number(r.OtherAmount) || 0;
    var promo = Number(r.PromoAmount) || 0;
    var medicine = (Number(r.MedicineAmount) || 0) + (Number(r.OralMedAmount) || 0); // ซื้อยาเพิ่ม (คิดค่ายา) + จ่ายยาทาน
    var amount = Number(r.Amount) || 0;
    var collected = round2_(amount + promo + canula + medicine + other);
    totalCollected += collected;
    canulaTotal += canula;
    otherAmountTotal += other;
    promoAmountTotal += promo;
    medicineAmountTotal += medicine;
    clinicNet += recordClinicShare_(r);
    if (r.PaymentMethod === 'เงินสด') cashAmount += collected;
    else if (r.PaymentMethod === 'โอนเงิน') transferAmount += collected;
    else if (r.PaymentMethod === 'บัตรเครดิต') cardAmount += collected;
    recordDoctorShares_(r).forEach(function (d) {
      if (!doctorMap[d.doctorId]) doctorMap[d.doctorId] = { name: d.name, amount: 0, topped: false };
      doctorMap[d.doctorId].amount += Number(d.share) || 0;
      // DoctorShareNote is only ever set by recomputeDoctorDayShares_ when a
      // record's share was topped up to meet the doctor's daily minimum —
      // surfacing that here lets the summary mark the doctor with a footnote
      // instead of silently showing an inflated number with no explanation.
      if (d.note) doctorMap[d.doctorId].topped = true;
    });
    if (r.SalesID) {
      if (!salesMap[r.SalesID]) salesMap[r.SalesID] = { name: r.SalesName, amount: 0 };
      salesMap[r.SalesID].amount += Number(r.SalesShare) || 0;
    }
    // Round 15 item 3: itemize down to each individual line item (its own
    // "ชื่อรายการ" name/label), not one aggregate row per record — the day
    // panel's "จ่าย supplier" list previously only showed the record's
    // patient name with no way to tell which product/medicine was bought.
    if (medicine > 0) {
      parseLineItems_(r.OralMedJSON).forEach(function (line) {
        var amt = Number(line.amount) || 0;
        if (amt > 0) medicineItems.push({ recordId: r.RID, patientName: r.PatientName, name: 'ยาทาน ' + (line.name || ''), amount: round2_(amt) });
      });
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
    out.hn = patient ? (patient.HN || '') : '';
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

/** A MedicineJSON line of รายการซื้อยาเพิ่มหน้าคลีนิค (คิดค่ายา): flagged charged, or (rows from before the flag) carrying an amount. */
function isChargedMedLine_(l) { return !!l && (l.charged === true || Number(l.amount) > 0); }

/** MedicineJSON lines -> {uncharged, charged}, the two groups of the record form (ไม่คิดค่ายา / คิดค่ายา). */
function splitMedicineLines_(r) {
  var all = parseLineItems_(r.MedicineJSON), out = { uncharged: [], charged: [] };
  all.forEach(function (l) { (isChargedMedLine_(l) ? out.charged : out.uncharged).push(l); });
  return out;
}

/**
 * รายการซื้อยาเพิ่มหน้าคลีนิค lines -> {items, total}; see the MedicineJSON note above. Both groups of the form come
 * through here and are stored together: (ไม่คิดค่ายา) lines — record.useMedicine / medicineItems — are taken out of Stock and
 * NOT charged (unitPrice/amount 0); (คิดค่ายา) lines — record.useChargedMed / chargedMedItems — are taken out of Stock too and
 * charged at qtyDispensed × unitPrice (flagged charged:true). Each keeps a unitCost snapshot (ราคาทุน, for รายงานรายเซลล์ —
 * never sent to the client, see mapRecordOut_). A line saved before the dispense table existed ({name, amount}) keeps what it
 * had if charged and is 0 otherwise. ctx: {stockById, oldLines}.
 */
function medicineLines_(record, ctx) {
  var out = { items: [], total: 0 };
  [
    { use: record.useMedicine, items: record.medicineItems, charged: false, label: 'ซื้อยาเพิ่ม ไม่คิดค่ายา' },
    { use: record.useChargedMed, items: record.chargedMedItems, charged: true, label: 'ซื้อยาเพิ่ม คิดค่ายา' }
  ].forEach(function (g) {
    if (!g.use || !Array.isArray(g.items)) return;
    g.items.forEach(function (it, i) {
      if (!it || (!String(it.name || '').trim() && !it.itemId)) return;
      var line = ' (' + g.label + ' บรรทัดที่ ' + (i + 1) + ')';
      if (!('qtyDispensed' in it)) { // pre-dispense-table line: name + amount only
        var legacy = g.charged ? requireNumber_(it.amount || 0, 'จำนวนเงิน' + line, { min: 0 }) : 0;
        out.items.push(g.charged ? { name: it.name || '', amount: legacy, charged: true } : { name: it.name || '', amount: 0 });
        out.total += legacy;
        return;
      }
      if (!it.itemId || !ctx.stockById[it.itemId]) throw ValidationError_('กรุณาเลือกยาจากรายการสต๊อกยา' + line);
      var old = (ctx.oldLines || []).filter(function (l) { return l.itemId === it.itemId && l.unitCost != null; })[0];
      var o = {
        itemId: it.itemId, name: String(it.name || '').trim(),
        qtyDispensed: requireNumber_(it.qtyDispensed || 0, 'จำนวนที่เบิก' + line, { min: 0 }),
        unitCost: old ? Number(old.unitCost) || 0 : Number(ctx.stockById[it.itemId].CostPrice) || 0
      };
      depositSplit_(o, it, line);
      if (g.charged) {
        o.charged = true;
        o.unitPrice = requireNumber_(it.unitPrice || 0, 'ราคาต่อหน่วย' + line, { min: 0 });
        o.amount = round2_(o.qtyDispensed * o.unitPrice);
      } else {
        o.unitPrice = 0;
        o.amount = 0;
      }
      out.items.push(o);
      out.total += o.amount;
    });
  });
  out.total = round2_(out.total);
  return out;
}

/**
 * รายการจ่ายยาทาน lines -> {items, total}: each is picked from the รายการยาทาน stock (itemId; name and unit come from the stock,
 * not the client), qty × ราคา/หน่วย is charged and qty leaves Stock (see recordStockEffects_). ctx: {stockById}.
 */
function oralMedLines_(record, ctx) {
  var out = { items: [], total: 0 };
  if (!record.useOralMed || !Array.isArray(record.oralMedItems)) return out;
  record.oralMedItems.forEach(function (it, i) {
    if (!it || (!String(it.name || '').trim() && !it.itemId)) return;
    var line = ' (จ่ายยาทาน บรรทัดที่ ' + (i + 1) + ')';
    var item = it.itemId ? ctx.stockById[it.itemId] : null;
    if (!item || stockTypeOf_(item) !== 'oral') throw ValidationError_('กรุณาเลือกยาจากรายการยาทาน' + line);
    var qty = requireNumber_(it.qty || 0, 'จำนวน' + line, { min: 0 });
    if (!qty) throw ValidationError_('จำนวนต้องมากกว่า 0' + line);
    var unitPrice = requireNumber_(it.unitPrice || 0, 'ราคา/หน่วย' + line, { min: 0 });
    var o = { itemId: item.ItemID, name: item.Name, qty: qty, unit: item.Unit || '', unitPrice: unitPrice, amount: round2_(qty * unitPrice) };
    out.items.push(o);
    out.total += o.amount;
  });
  out.total = round2_(out.total);
  return out;
}

/** รายการ Canula (เข็มทู่) lines {name, qty, amount} -> {items, total}; a client without the lines sends only canulaAmount (one line). */
function canulaLines_(record) {
  if (!record.useCanula) return { items: [], total: 0 };
  var src = Array.isArray(record.canulaItems) ? record.canulaItems.filter(function (it) { return it && (String(it.name || '').trim() || Number(it.amount) > 0); }) : null;
  if (!src) src = Number(record.canulaAmount) > 0 ? [{ name: '', qty: 1, amount: record.canulaAmount }] : [];
  var total = 0;
  var items = src.map(function (it, i) {
    var line = ' (Canula บรรทัดที่ ' + (i + 1) + ')';
    var amount = requireNumber_(it.amount || 0, 'จำนวนเงิน' + line, { min: 0 });
    total += amount;
    return { name: String(it.name || '').trim(), qty: requireNumber_(it.qty == null || it.qty === '' ? 1 : it.qty, 'จำนวนที่ใช้' + line, { min: 0 }), amount: amount };
  });
  return { items: items, total: round2_(total) };
}

/**
 * รายการยาจากใบเบิก -> the WithdrawJSON lines. Each names a จัดยา-เบิกยา line
 * (withdrawId) that must be this patient's and not taken in by another
 * record; its item and qtyDispensed come from that row (it already left
 * Stock). Not charged — only the used / ยาฝาก split is the client's.
 * ctx: {rid, patientId, wdById}.
 */
function withdrawLines_(record, ctx) {
  var seen = {};
  if (record.useWithdraw === false) return []; // unticked: the lines go back to waiting
  return (Array.isArray(record.withdrawItems) ? record.withdrawItems : []).filter(function (it) { return it && it.withdrawId; }).map(function (it, i) {
    var line = ' (ยาจากใบเบิก บรรทัดที่ ' + (i + 1) + ')';
    var wd = ctx.wdById[it.withdrawId];
    if (!wd || wd.PatientID !== ctx.patientId) throw ValidationError_('ไม่พบรายการเบิกยาของคนไข้รายนี้ (อาจถูกลบไปแล้ว)' + line);
    if (wd.RecordID && wd.RecordID !== ctx.rid) throw ValidationError_('รายการเบิกยา ' + wd.ItemName + ' ถูกบันทึกในการรักษา ' + wd.RecordID + ' แล้ว' + line);
    if (seen[it.withdrawId]) throw ValidationError_('รายการเบิกยา ' + wd.ItemName + ' ซ้ำ' + line);
    seen[it.withdrawId] = true;
    var o = { withdrawId: wd.WithdrawID, itemId: wd.ItemID, name: wd.ItemName, qtyDispensed: Number(wd.Qty) || 0 };
    return depositSplit_(o, it, line);
  });
}

/** Sets qtyUsed / qtyDeposit (= dispensed − used, becomes a ยาฝาก lot) / its expiry on line o from client line it. */
function depositSplit_(o, it, line) {
  var qu = requireNumber_(it.qtyUsed || 0, 'จำนวนที่ใช้' + line, { min: 0 });
  if (qu > o.qtyDispensed) throw ValidationError_('จำนวนที่ใช้มากกว่าจำนวนที่เบิก' + line);
  o.qtyUsed = qu;
  o.qtyDeposit = round2_(o.qtyDispensed - qu);
  o.expiry = o.qtyDeposit > 0 && it.expiry ? String(it.expiry) : '';
  if (o.expiry) requireDateStr_(o.expiry, 'วันหมดอายุยาฝาก' + line);
  return o;
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
 * What a saved Records row takes out of Stock: {itemId: qty}. Its ยาจากใบเบิก
 * (WithdrawJSON) took nothing — the withdrawal already did. Rows without StockApplied=TRUE (saved before this existed)
 * take nothing.
 */
function recordStockEffects_(r) {
  var eff = {};
  if (!isTrue_(r.StockApplied)) return eff;
  var add = function (key, n) { if (key && n) eff[key] = round2_((eff[key] || 0) + n); };
  if (isTrue_(r.UseSupplies)) {
    parseLineItems_(r.SuppliesJSON).forEach(function (l) { add(l.itemId, Number(l.qty) || 0); });
  }
  if (isTrue_(r.UseMedicine)) {
    parseLineItems_(r.MedicineJSON).forEach(function (l) { add(l.itemId, Number(l.qtyDispensed) || 0); });
  }
  if (isTrue_(r.UseOralMed)) {
    parseLineItems_(r.OralMedJSON).forEach(function (l) { add(l.itemId, Number(l.qty) || 0); });
  }
  return eff;
}

/**
 * Moves Stock by (newEff - oldEff) — so a new record applies its lines, an
 * edit applies only what changed, and a delete (newEff empty) puts
 * everything back. Stock may go negative; only a missing item throws (before
 * anything is written). Moves are StockLog rows of kind `logKind` (default
 * 'record'; จัดยา-เบิกยา passes 'withdraw') dated today — not the record's
 * date, so they are always the item's latest entry and Stock.Qty follows
 * them. Must run inside the caller's script lock.
 */
function applyStockEffects_(session, oldEff, newEff, label, logKind) {
  var stockDelta = {}; // + = more taken out of stock
  Object.keys(oldEff).concat(Object.keys(newEff)).forEach(function (k) { var d = round2_((newEff[k] || 0) - (oldEff[k] || 0)); if (d) stockDelta[k] = d; });
  if (!Object.keys(stockDelta).length) return;
  var stockById = {};
  getRows_(getSheet_(CONFIG.SHEETS.STOCK)).forEach(function (r) { stockById[r.ItemID] = r; });
  Object.keys(stockDelta).forEach(function (id) {
    if (!stockById[id] && stockDelta[id] > 0) throw ValidationError_('ไม่พบรายการในสต๊อก (อาจถูกลบไปแล้ว)');
  });
  var logSheet = getSheet_(CONFIG.SHEETS.STOCK_LOG);
  var movedIds = Object.keys(stockDelta).filter(function (id) { return stockById[id]; });
  movedIds.forEach(function (id) {
    var item = stockById[id], d = stockDelta[id], cur = Number(item.Qty) || 0;
    var left = round2_(cur - d);
    appendRow_(logSheet, CONFIG.HEADERS.StockLog,
      stockLogRow_(session, 'SL' + nextSeq_('STOCKLOG', 5), todayStr_(), item, logKind || 'record', d < 0 ? -d : 0, d > 0 ? d : 0, left, label));
    // The log row just written IS the latest balance, so set Stock.Qty from it (resyncStockQtyFromLog_ would read the
    // whole, ever-growing StockLog again for every item).
    updateRowByIndex_(getSheet_(CONFIG.SHEETS.STOCK), CONFIG.HEADERS.Stock, item._row, Object.assign({}, item, { Qty: left, UpdatedAt: nowStr_() }));
    stockById[id] = Object.assign({}, item, { Qty: left });
  });
  if (movedIds.length) invalidateStockLiteCache_();
}

/* ---------------------------------------------------------------------
   ยาฝาก lots (MedDeposits). A record's deposit side is:
     deposits: {'itemId|expiry': qty} — its รายการยา lines' จำนวนฝาก, kept
       as lots with RecordID = the record (see recordDepositSide_);
     uses: its รายการใช้ยาฝาก lines {itemId, name, qty, lots:[{depositId,
       qty}]} — `lots` says which lots each line was taken from, so an edit
       or delete can put it back exactly (older lines have none and go back
       to the patient's lot of that item without a record).
   --------------------------------------------------------------------- */
function recordDepositSide_(r) {
  var side = { deposits: {}, uses: [] };
  if (!isTrue_(r.StockApplied)) return side;
  (isTrue_(r.UseMedicine) ? parseLineItems_(r.MedicineJSON) : []).concat(parseLineItems_(r.WithdrawJSON)).forEach(function (l) {
    var q = Number(l.qtyDeposit) || 0;
    if (l.itemId && q > 0) { var k = l.itemId + '|' + (l.expiry || ''); side.deposits[k] = round2_((side.deposits[k] || 0) + q); }
  });
  if (isTrue_(r.UseDepositUse)) side.uses = parseLineItems_(r.DepositUseJSON).filter(function (l) { return l.itemId; });
  return side;
}

/** A lot is expired once its ExpiryDate is before today; blank never expires. */
function depositExpired_(lot) {
  var e = dateStr_(lot.ExpiryDate);
  return !!e && e < todayStr_();
}

/**
 * Works out every MedDeposits change for moving one record's deposit side
 * from `oldSide` to `newSide`, in memory only — throws (nothing written yet)
 * if a lot would go below zero or the patient lacks enough unexpired ยาฝาก.
 * Order: put back what the old version used, adjust the lots it deposited,
 * then take the new uses from the earliest-expiring unexpired lots (blank
 * expiry last). Returns {uses, commit}: `uses` are newSide.uses with their
 * `lots` filled in (store them on the record); commit() writes the rows.
 * ctx: {rid, date, patientId, patientName}. Must run inside the script lock.
 */
function planDeposits_(session, ctx, oldSide, newSide) {
  var sheet = getSheet_(CONFIG.SHEETS.MED_DEPOSITS);
  var lots = getRows_(sheet).filter(function (r) { return r.PatientID === ctx.patientId; });
  var byId = {}, changed = {}, added = [];
  lots.forEach(function (l) { byId[l.DepositID] = l; });
  var itemNames = null;
  var itemName = function (itemId) {
    if (!itemNames) { itemNames = {}; getRows_(getSheet_(CONFIG.SHEETS.STOCK)).forEach(function (s) { itemNames[s.ItemID] = s.Name; }); }
    return itemNames[itemId] || '';
  };
  var newLot = function (itemId, expiry, recordId, qty) {
    var l = {
      DepositID: 'DP' + nextSeq_('DEPOSIT', 5), Date: ctx.date || todayStr_(), PatientID: ctx.patientId, PatientName: ctx.patientName || '',
      ItemName: itemName(itemId), QtyRemaining: 0, CreatedBy: session.username, ItemID: itemId, UpdatedAt: nowStr_(),
      ExpiryDate: expiry || '', RecordID: recordId || '', QtyDeposited: 0
    };
    lots.push(l); byId[l.DepositID] = l; added.push(l);
    return l;
  };
  var move = function (l, qty) { l.QtyRemaining = round2_((Number(l.QtyRemaining) || 0) + qty); changed[l.DepositID] = l; };
  // The patient's lot of an item that isn't tied to a record (the pre-lot balance rows).
  var looseLot = function (itemId) {
    return lots.filter(function (l) { return l.ItemID === itemId && !l.RecordID; })[0] ||
      lots.filter(function (l) { return l.ItemID === itemId; })[0] || newLot(itemId, '', '', 0);
  };

  // 1. Put back what the old version used.
  oldSide.uses.forEach(function (u) {
    if (Array.isArray(u.lots) && u.lots.length) {
      u.lots.forEach(function (p) { move(byId[p.depositId] || looseLot(u.itemId), Number(p.qty) || 0); });
    } else {
      move(looseLot(u.itemId), Number(u.qty) || 0);
    }
  });

  // 2. Lots this record deposited: change each by (new - old).
  Object.keys(oldSide.deposits).concat(Object.keys(newSide.deposits)).forEach(function (k, i, all) {
    if (all.indexOf(k) !== i) return;
    var d = round2_((newSide.deposits[k] || 0) - (oldSide.deposits[k] || 0));
    if (!d) return;
    var parts = k.split('|'), itemId = parts[0], expiry = parts[1];
    var lot = lots.filter(function (l) { return l.RecordID === ctx.rid && l.ItemID === itemId && dateStr_(l.ExpiryDate) === expiry; })[0];
    if (!lot && oldSide.deposits[k]) lot = looseLot(itemId); // deposited before lots existed
    if (!lot) lot = newLot(itemId, expiry, ctx.rid, 0);
    move(lot, d);
    if (lot.RecordID === ctx.rid) lot.QtyDeposited = round2_((Number(lot.QtyDeposited) || 0) + d);
    if (lot.QtyRemaining < 0) throw ValidationError_('ยาฝาก ' + (lot.ItemName || itemName(itemId)) + ' จากรายการนี้ถูกใช้ไปแล้ว ลดจำนวนฝากไม่ได้');
  });

  // 3. Take the new uses, earliest expiry first.
  var uses = newSide.uses.map(function (u) {
    var need = Number(u.qty) || 0, taken = [];
    var avail = lots.filter(function (l) { return l.ItemID === u.itemId && (Number(l.QtyRemaining) || 0) > 0 && !depositExpired_(l); })
      .sort(function (a, b) {
        var ea = dateStr_(a.ExpiryDate) || '9999', eb = dateStr_(b.ExpiryDate) || '9999';
        return ea.localeCompare(eb) || dateStr_(a.Date).localeCompare(dateStr_(b.Date));
      });
    avail.forEach(function (l) {
      if (need <= 0) return;
      var q = Math.min(need, Number(l.QtyRemaining) || 0);
      move(l, -q);
      taken.push({ depositId: l.DepositID, qty: round2_(q) });
      need = round2_(need - q);
    });
    if (need > 0) {
      var left = round2_((Number(u.qty) || 0) - need);
      throw ValidationError_('ยาฝาก ' + (u.name || itemName(u.itemId)) + ' ไม่พอ (คงเหลือที่ยังไม่หมดอายุ ' + left + ')');
    }
    return Object.assign({}, u, { lots: taken });
  });

  // Every touched lot must still be >= 0 (step 1/2 can only add, except a reduced deposit, checked above).
  Object.keys(changed).forEach(function (id) {
    if ((Number(changed[id].QtyRemaining) || 0) < 0) throw ValidationError_('ยาฝาก ' + changed[id].ItemName + ' ไม่พอ');
  });

  return {
    uses: uses,
    commit: function () {
      Object.keys(changed).forEach(function (id) {
        var l = changed[id];
        l.CreatedBy = session.username; l.UpdatedAt = nowStr_();
        if (added.indexOf(l) > -1) appendRow_(sheet, CONFIG.HEADERS.MedDeposits, l);
        else updateRowByIndex_(sheet, CONFIG.HEADERS.MedDeposits, l._row, l);
      });
    }
  };
}

function mapDepositOut_(r) {
  return {
    id: r.DepositID, itemId: r.ItemID, name: r.ItemName, qty: Number(r.QtyRemaining) || 0,
    qtyDeposited: numOrNull_(r.QtyDeposited), date: dateStr_(r.Date), expiry: dateStr_(r.ExpiryDate),
    expired: depositExpired_(r), recordId: r.RecordID || '', createdBy: r.CreatedBy
  };
}

/** ยาฝาก lots the patient still has (balance > 0) — record modal and patient panel. */
function patientDeposits_(patientId) {
  return getRows_(getSheet_(CONFIG.SHEETS.MED_DEPOSITS)).filter(function (r) { return r.PatientID === patientId && (Number(r.QtyRemaining) || 0) > 0; })
    .map(mapDepositOut_)
    .sort(function (a, b) { return (a.expiry || '9999').localeCompare(b.expiry || '9999') || a.date.localeCompare(b.date); });
}
function listPatientDeposits(token, patientId) {
  return safeCall_(function () {
    requireAuth_(token, 'staff');
    return { ok: true, data: patientDeposits_(patientId) };
  });
}

/** เมนู "รายการยาฝาก": every lot still holding ยา (balance > 0), with the patient's HN and the item's unit; the page filters/pages it. */
function listDeposits(token) {
  return safeCall_(function () {
    requireAuth_(token, 'staff');
    var hnById = {}, unitById = {};
    getRows_(getSheet_(CONFIG.SHEETS.PATIENTS)).forEach(function (p) { hnById[p.ID] = p.HN || ''; });
    getRows_(getSheet_(CONFIG.SHEETS.STOCK)).forEach(function (s) { unitById[s.ItemID] = s.Unit || ''; });
    var rows = getRows_(getSheet_(CONFIG.SHEETS.MED_DEPOSITS)).filter(function (r) { return (Number(r.QtyRemaining) || 0) > 0; });
    var out = rows.map(function (r) {
      return Object.assign(mapDepositOut_(r), { patientId: r.PatientID, patientName: r.PatientName, hn: hnById[r.PatientID] || '', unit: unitById[r.ItemID] || '' });
    });
    out.sort(function (a, b) { return (a.expiry || '9999').localeCompare(b.expiry || '9999') || a.patientName.localeCompare(b.patientName, 'th'); });
    return { ok: true, data: out };
  });
}

/**
 * The movement of one ยาฝาก lot for the side panel: a summary and the lines
 * {date, label, recordId, delta, balance} — the deposit (+), each use by a
 * บันทึกการรักษา (−, from the lots its DepositUseJSON names) and, if the
 * balance was corrected by an admin, a ปรับยอด line for the difference.
 */
function getDepositMovements(token, id) {
  return safeCall_(function () {
    requireAuth_(token, 'staff');
    var lot = getRows_(getSheet_(CONFIG.SHEETS.MED_DEPOSITS)).filter(function (r) { return r.DepositID === id; })[0];
    if (!lot) throw ValidationError_('ไม่พบรายการยาฝาก');
    var patient = getRows_(getSheet_(CONFIG.SHEETS.PATIENTS)).filter(function (p) { return p.ID === lot.PatientID; })[0];
    var stockRow = getRows_(getSheet_(CONFIG.SHEETS.STOCK)).filter(function (s) { return s.ItemID === lot.ItemID; })[0];
    var moves = [], used = 0;
    getRows_(getSheet_(CONFIG.SHEETS.RECORDS)).forEach(function (r) {
      parseLineItems_(r.DepositUseJSON).forEach(function (u) {
        (u.lots || []).forEach(function (p) {
          var q = Number(p.qty) || 0;
          if (p.depositId === id && q) { used = round2_(used + q); moves.push({ date: dateStr_(r.Date), sort: '1' + String(r.CreatedAt), label: 'ใช้ยาฝาก', recordId: r.RID, delta: -q }); }
        });
      });
    });
    var remaining = Number(lot.QtyRemaining) || 0;
    var deposited = numOrNull_(lot.QtyDeposited);
    if (deposited == null) deposited = round2_(remaining + used); // lot from before QtyDeposited was kept
    moves.push({ date: dateStr_(lot.Date), sort: '0', label: 'ฝากยา', recordId: lot.RecordID || '', delta: deposited });
    var adjust = round2_(remaining - (deposited - used));
    if (adjust) moves.push({ date: dateStr_(lot.UpdatedAt) || dateStr_(lot.Date), sort: '2', label: 'ปรับยอดโดยผู้ดูแล', recordId: '', delta: adjust });
    moves.sort(function (a, b) { return a.date.localeCompare(b.date) || a.sort.localeCompare(b.sort); });
    var bal = 0;
    moves.forEach(function (m) { bal = round2_(bal + m.delta); m.balance = bal; delete m.sort; });
    var summary = Object.assign(mapDepositOut_(lot), {
      patientId: lot.PatientID, patientName: lot.PatientName, hn: patient ? (patient.HN || '') : '', unit: stockRow ? (stockRow.Unit || '') : '',
      deposited: deposited, used: used
    });
    return { ok: true, data: { lot: summary, movements: moves } };
  });
}

/**
 * Edits one ยาฝาก lot: {id, expiry, qty}. Anyone may set the วันหมดอายุ;
 * changing the จำนวนคงเหลือ (a correction after a count, or writing off an
 * expired lot by setting 0) is admin-only.
 */
function saveDeposit(token, payload) {
  return safeCall_(function () {
    var session = requireAuth_(token, 'staff');
    payload = payload || {};
    if (payload.expiry) requireDateStr_(payload.expiry, 'วันหมดอายุ');
    var lock = LockService.getScriptLock();
    lock.waitLock(30000);
    try {
      var sheet = getSheet_(CONFIG.SHEETS.MED_DEPOSITS);
      var row = getRows_(sheet).filter(function (r) { return r.DepositID === payload.id; })[0];
      if (!row) throw ValidationError_('ไม่พบรายการยาฝาก');
      var upd = { ExpiryDate: payload.expiry || '', UpdatedAt: nowStr_(), CreatedBy: session.username };
      var oldQty = Number(row.QtyRemaining) || 0, detail = 'expiry ' + (payload.expiry || '-');
      if (payload.qty !== undefined && payload.qty !== '' && Number(payload.qty) !== oldQty) {
        if (session.level !== 'admin') throw ValidationError_('การแก้ไขจำนวนยาฝากต้องใช้สิทธิ์ผู้ดูแลระบบ');
        upd.QtyRemaining = requireNumber_(payload.qty, 'จำนวนคงเหลือ', { min: 0 });
        detail += ' · qty ' + oldQty + ' -> ' + upd.QtyRemaining;
      }
      updateRowByIndex_(sheet, CONFIG.HEADERS.MedDeposits, row._row, Object.assign({}, row, upd));
      writeAudit_(session, 'UPDATE', 'MedDeposit', row.DepositID, row.PatientName + ' ' + row.ItemName + ' · ' + detail);
      return { ok: true };
    } finally {
      lock.releaseLock();
    }
  });
}

function saveRecord(token, record) {
  return safeCall_(function () {
    var session = requireAuth_(token, 'staff');
    // Round 14 item 1.1: หมวดหมู่การรักษา field removed from the modal UI —
    // no longer a required field. ServiceCategory column stays in the
    // schema (append-only discipline) and still stores whatever the client
    // sends (now always '' going forward); old records keep their value.
    // วิธีชำระเงิน/สถานะการชำระ are set afterwards by payRecord (ขั้นตอนชำระเงิน).
    // แพทย์ is chosen per หัตถการ/โปร line (one visit may pay several doctors).
    requireFields_(record, { date: 'วันที่', patientId: 'คนไข้' });
    requireDateStr_(record.date, 'วันที่');
    // "amount" here is จำนวนเงินหัตถการ — the base procedure price that
    // sales/doctor commission % are computed against: the sum of the
    // รายการหัตถการ lines (a client without them still sends a plain amount).
    // Canula/promo/medicine/other money are tracked separately (see header
    // schema comment) — only amount (always) and promo (round 13) feed the
    // commissionable base.
    var procedureItems = (record.useProcedure && Array.isArray(record.procedureItems) ? record.procedureItems : [])
      .filter(function (it) { return it && (it.procId || String(it.name || '').trim()); });
    var amount = Array.isArray(record.procedureItems)
      ? round2_(procedureItems.reduce(function (s, it, i) { return s + requireNumber_(it.amount || 0, 'จำนวนเงินหัตถการ (บรรทัดที่ ' + (i + 1) + ')', { min: 0 }); }, 0))
      : requireNumber_(record.amount || 0, 'จำนวนเงินหัตถการ', { min: 0 });
    var canulaGroup = canulaLines_(record);
    var canulaAmount = canulaGroup.total;
    // รายการเวชภัณฑ์ที่ใช้ — quantity only (no money); blank-name lines dropped;
    // lines picked from stock (itemId) are deducted via applyStockEffects_.
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
    // รายการยา (medicineLines_) needs the sheets, so it is built inside the lock.
    var ownMedItems = qtyLines_(record.ownMedItems, !!record.useOwnMed, 'ถือยาเข้า', [], false);
    var depositUseItems = qtyLines_(record.depositUseItems, !!record.useDepositUse, 'ใช้ยาฝาก', ['itemId'], true);
    var messengerItems = qtyLines_(record.messengerItems, !!record.useMessenger, 'Messenger/พัสดุ', ['note'], false);
    var otherGroup = sumLineItems_(record.otherItems, !!record.useOther, 'จำนวนเงินอื่นๆ');
    var otherAmount = otherGroup.total;

    var lock = LockService.getScriptLock();
    lock.waitLock(30000);
    beginWriteScope_();
    try {
      var sheet = getSheet_(CONFIG.SHEETS.RECORDS);
      var isNew = !record.id;
      var patientRows = getRows_(getSheet_(CONFIG.SHEETS.PATIENTS));
      var patient = patientRows.filter(function (p) { return p.ID === record.patientId; })[0];
      if (!patient) throw ValidationError_('ไม่พบข้อมูลคนไข้ที่เลือก');

      var doctorById = {};
      getRows_(getSheet_(CONFIG.SHEETS.DOCTORS)).forEach(function (d) { doctorById[d.DoctorID] = d; });
      // Every หัตถการ/โปร line names its doctor (legacy clients send one record-level doctorId).
      var lineDoctor = function (line, label) {
        var id = line.doctorId || record.doctorId;
        if (!id) throw ValidationError_('กรุณาเลือกแพทย์ (' + label + ')');
        if (!doctorById[id]) throw ValidationError_('ไม่พบข้อมูลแพทย์ที่เลือก (' + label + ')');
        return id;
      };
      var doctorEntries = [];

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
      var promoAmount = 0, promoOut = [];
      if (record.usePromo && Array.isArray(record.promoItems) && record.promoItems.length) {
        var allPromos = getRows_(getSheet_(CONFIG.SHEETS.PROMOTIONS));
        record.promoItems.forEach(function (line, i) {
          var promo = allPromos.filter(function (p) { return p.PromoID === line.promoId; })[0];
          if (!promo) throw ValidationError_('ไม่พบโปรโมชันที่เลือก (อาจถูกลบไปแล้ว)');
          var lineAmount = requireNumber_(line.amount || 0, 'จำนวนเงินโปร', { min: 0 });
          var doctorPct = Number(promo.DoctorSharePct) || 0;
          var doctorId = lineDoctor(line, 'โปร บรรทัดที่ ' + (i + 1));
          promoAmount += lineAmount;
          doctorEntries.push({ doctorId: doctorId, extra: lineAmount * doctorPct / 100 });
          promoOut.push({ promoId: promo.PromoID, name: promo.Name, price: Number(promo.Price) || 0, doctorSharePct: doctorPct, amount: lineAmount, doctorId: doctorId });
        });
      }
      promoAmount = round2_(promoAmount);

      if (!isNew) {
        var existingRows = getRows_(sheet);
        var existing = existingRows.filter(function (r) { return r.RID === record.id; })[0];
        if (!existing) throw ValidationError_('ไม่พบข้อมูลการรักษาที่จะแก้ไข');
        // Editing a record keeps the sales % it was saved with; only picking a
        // different sales person takes that person's current rate.
        if (sales && existing.SalesID === sales.SalesID && existing.SalesCommissionPct !== '' && existing.SalesCommissionPct != null) {
          sales = Object.assign({}, sales, { CommissionRate: Number(existing.SalesCommissionPct) || 0 });
        }
      }
      var rid = isNew ? ('R' + nextSeq_('RECORD', 6)) : record.id;

      // รายการหัตถการ — name/price re-read from the Procedures sheet; only the
      // line's amount (pre-filled from the price, editable) is the client's.
      // A line whose procedure was since deleted keeps its saved name.
      var procById = {};
      getRows_(getSheet_(CONFIG.SHEETS.PROCEDURES)).forEach(function (p) { procById[p.ProcID] = p; });
      var procedureOut = procedureItems.map(function (it, i) {
        var p = procById[it.procId];
        if (!p && !String(it.name || '').trim()) throw ValidationError_('ไม่พบหัตถการที่เลือก (อาจถูกลบไปแล้ว)');
        var doctorId = lineDoctor(it, 'หัตถการ บรรทัดที่ ' + (i + 1));
        doctorEntries.push({ doctorId: doctorId, base: Number(it.amount) || 0 });
        return { procId: p ? p.ProcID : (it.procId || ''), name: p ? p.Name : String(it.name).trim(), price: p ? Number(p.Price) || 0 : Number(it.price) || 0, amount: Number(it.amount) || 0, doctorId: doctorId, detail: String(it.detail || '').trim().slice(0, 500) };
      });

      var wdSheet = getSheet_(CONFIG.SHEETS.MED_WITHDRAWALS);
      var wdRows = getRows_(wdSheet), wdById = {}, stockById = {};
      wdRows.forEach(function (w) { wdById[w.WithdrawID] = w; });
      getRows_(getSheet_(CONFIG.SHEETS.STOCK)).forEach(function (s) { stockById[s.ItemID] = s; });
      var medicineGroup = medicineLines_(record, { stockById: stockById, oldLines: isNew ? [] : parseLineItems_(existing.MedicineJSON) });
      var withdrawItems = withdrawLines_(record, { rid: rid, patientId: patient.ID, wdById: wdById });
      var oralGroup = oralMedLines_(record, { stockById: stockById });

      var salesPct = sales ? (Number(sales.CommissionRate) || 0) : 0;
      var doctorShares = buildDoctorShares_(doctorEntries, doctorById);
      // สถานะการรักษา: a save keeps the record's step; "บันทึกการรักษา" (record.complete) moves it on to
      // รอชำระเงิน; a record made by hand (no จัดยา-เบิกยา first) starts at พบแพทย์. Paying makes it เสร็จสิ้น.
      var visitFields = isNew
        ? { VisitStatus: 'dispensed', QueueDate: '', QueueDoctor: '', QueuePay: '', VisitDoctorID: '', VisitDoctorName: '' }
        : { VisitStatus: visitStatusOf_(existing), QueueDate: existing.QueueDate || '', QueueDoctor: existing.QueueDoctor || '', QueuePay: existing.QueuePay || '', VisitDoctorID: existing.VisitDoctorID || '', VisitDoctorName: existing.VisitDoctorName || '' };

      var obj = {
        RID: rid,
        Date: record.date, PatientID: patient.ID, PatientName: patientFullName_(patient),
        Phone: patient.Phone, ServiceCategory: record.serviceCategory || '', Detail: record.detail || '',
        SalesID: record.salesId || '', SalesName: salesName,
        // Payment fields only change through payRecord.
        PaymentMethod: isNew ? '' : existing.PaymentMethod, SuppliesJSON: JSON.stringify(supplies),
        UseCanula: record.useCanula ? 'TRUE' : 'FALSE', CanulaAmount: record.useCanula ? canulaAmount : 0, CanulaJSON: JSON.stringify(canulaGroup.items),
        Amount: amount, Status: isNew ? 'unpaid' : (existing.Status || 'unpaid'), ReceiptNo: isNew ? '' : (existing.ReceiptNo || ''),
        Note: record.note || '', CreatedBy: isNew ? session.username : (existing.CreatedBy || session.username),
        CreatedAt: isNew ? nowStr_() : (existing.CreatedAt || nowStr_()), UpdatedAt: nowStr_(),
        OtherAmount: otherAmount,
        SalesCommissionPct: salesPct, SalesShare: round2_(amount * salesPct / 100),
        VisitStatus: visitFields.VisitStatus, QueueDate: visitFields.QueueDate, QueueDoctor: visitFields.QueueDoctor, QueuePay: visitFields.QueuePay, VisitDoctorID: visitFields.VisitDoctorID, VisitDoctorName: visitFields.VisitDoctorName,
        UsePromo: record.usePromo ? 'TRUE' : 'FALSE', PromoJSON: JSON.stringify(promoOut), PromoAmount: promoAmount,
        UseMedicine: medicineGroup.items.length ? 'TRUE' : 'FALSE', MedicineJSON: JSON.stringify(medicineGroup.items), MedicineAmount: medicineGroup.total,
        UseOralMed: oralGroup.items.length ? 'TRUE' : 'FALSE', OralMedJSON: JSON.stringify(oralGroup.items), OralMedAmount: oralGroup.total,
        UseOther: record.useOther ? 'TRUE' : 'FALSE', OtherItemsJSON: JSON.stringify(otherGroup.items),
        SupplierPayAmount: round2_(medicineGroup.total + otherAmount),
        // Round 14 item 1.7 — OPD photo attachment (Drive fileId, uploaded
        // client-side via the shared resumable-chunk-upload flow before
        // save; see startResumableUpload's 'opd-photo' kind).
        OpdImageFileId: record.opdImageFileId || '',
        UseSupplies: supplies.length ? 'TRUE' : 'FALSE',
        UseOwnMed: ownMedItems.length ? 'TRUE' : 'FALSE', OwnMedJSON: JSON.stringify(ownMedItems),
        UseDepositUse: depositUseItems.length ? 'TRUE' : 'FALSE', DepositUseJSON: JSON.stringify(depositUseItems),
        UseMessenger: messengerItems.length ? 'TRUE' : 'FALSE', MessengerJSON: JSON.stringify(messengerItems),
        StockApplied: isNew ? 'TRUE' : (existing.StockApplied || ''),
        UseProcedure: procedureOut.length ? 'TRUE' : 'FALSE', ProcedureJSON: JSON.stringify(procedureOut),
        WithdrawJSON: JSON.stringify(withdrawItems)
      };
      applyDoctorShares_(obj, doctorShares); // DoctorID/DoctorName/shares/ClinicShare
      if (isNew) {
        // Made by hand: the customer is already with the doctor of its first หัตถการ/โปร line (if any).
        var firstDoctor = doctorEntries.length ? doctorById[doctorEntries[0].doctorId] : null;
        applyVisitStatus_(obj, 'with_doctor', record.date, false, firstDoctor ? { id: firstDoctor.DoctorID, name: firstDoctor.Name } : { id: '', name: '' });
      }
      if (record.complete && visitStatusOf_(obj) !== 'completed') applyVisitStatus_(obj, 'waiting_payment', todayStr_(), false);
      obj.VisitStatus = visitStatusOf_(obj);

      // ยาฝาก + Stock before the record itself: both throw before writing
      // anything if a deposit lot would go negative / an item is missing.
      // The planned uses carry the lots they came from (see planDeposits_).
      var label = 'บันทึกการรักษา ' + obj.RID + ' · ' + obj.PatientName;
      var depositPlan = planDeposits_(session, { rid: rid, date: record.date, patientId: patient.ID, patientName: obj.PatientName },
        isNew ? { deposits: {}, uses: [] } : recordDepositSide_(existing), recordDepositSide_(obj));
      if (isTrue_(obj.StockApplied)) obj.DepositUseJSON = JSON.stringify(depositPlan.uses);
      applyStockEffects_(session, isNew ? {} : recordStockEffects_(existing), recordStockEffects_(obj), label);
      depositPlan.commit();
      linkWithdrawals_(wdSheet, wdRows, rid, withdrawItems);

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
      // above is done, and hand that single snapshot to every
      // recomputeDoctorDayShares_ call — appendRow_/updateRowByIndex_ never
      // shift existing row indices (only a delete does), and each call
      // updates the rows it rewrites in the snapshot itself. Every doctor the
      // record pays now or paid before (on the old date too, if it moved)
      // gets their daily-minimum top-up re-derived.
      var postWriteRows = getRows_(sheet);
      recomputeRecordDoctors_(isNew ? null : existing, obj, postWriteRows);

      writeAudit_(session, isNew ? 'CREATE' : 'UPDATE', 'Record', obj.RID, obj.PatientName + ' ' + obj.Amount);
      bumpDashboardCacheVersion_(); // Round 19 item 4: treatment records feed dashboard totals
      // The snapshot carries whatever top-up was just written onto this record.
      var savedRow = postWriteRows.filter(function (r) { return r.RID === obj.RID; })[0];
      return { ok: true, data: mapRecordOut_(savedRow || Object.assign({ _row: 0 }, obj)) };
    } finally {
      endWriteScope_();
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
      // (throws, deleting nothing, if its deposited medicine was already used),
      // and its จัดยา-เบิกยา lines go back to waiting for a record.
      var depositPlan = planDeposits_(session, { rid: id, date: dateStr_(existing.Date), patientId: existing.PatientID, patientName: existing.PatientName },
        recordDepositSide_(existing), { deposits: {}, uses: [] });
      applyStockEffects_(session, recordStockEffects_(existing), {}, 'ลบบันทึกการรักษา ' + id + ' · ' + existing.PatientName);
      depositPlan.commit();
      var wdSheet = getSheet_(CONFIG.SHEETS.MED_WITHDRAWALS);
      linkWithdrawals_(wdSheet, getRows_(wdSheet), id, []);
      deleteRowByIndex_(sheet, existing._row);
      // The deleted record's share of the daily minimum (if any) needs to
      // be re-distributed among whatever's left in each doctor's group.
      recomputeRecordDoctors_(existing, null);
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
/** The record's receipt number, assigning the next one the first time. */
function assignReceiptNo_(rid) {
  var sheet = getSheet_(CONFIG.SHEETS.RECORDS);
  var record = getRows_(sheet).filter(function (r) { return r.RID === rid; })[0];
  if (!record) throw ValidationError_('ไม่พบข้อมูลการรักษา');
  var receiptNo = record.ReceiptNo;
  if (!receiptNo) {
    receiptNo = generateReceiptNo_();
    updateRowByIndex_(sheet, CONFIG.HEADERS.Records, record._row, Object.assign({}, record, { ReceiptNo: receiptNo, UpdatedAt: nowStr_() }));
  }
  return receiptNo;
}

/**
 * The receipt as a complete HTML page (Receipt.html evaluated for this record)
 * for the in-app preview, which shows it in a frame with พิมพ์ / บันทึก PDF.
 * Assigns the receipt number on first use, like issueReceipt.
 */
function getReceiptHtml(token, rid) {
  return safeCall_(function () {
    var session = requireAuth_(token, 'staff');
    var receiptNo = assignReceiptNo_(rid);
    var tpl = HtmlService.createTemplateFromFile('Receipt');
    tpl.valid = true;
    tpl.data = buildReceiptData_(rid);
    writeAudit_(session, 'ISSUE_RECEIPT', 'Record', rid, receiptNo);
    return { ok: true, data: { receiptNo: receiptNo, html: tpl.evaluate().getContent() } };
  });
}

function issueReceipt(token, rid) {
  return safeCall_(function () {
    var session = requireAuth_(token, 'staff');
    var receiptNo = assignReceiptNo_(rid);

    var ticket = Utilities.getUuid();
    CacheService.getScriptCache().put(CONFIG.CACHE.RECEIPT_TICKET_PREFIX + ticket, JSON.stringify({ rid: rid }), CONFIG.CACHE.RECEIPT_TICKET_TTL);
    var url = ScriptApp.getService().getUrl() + '?page=receipt&ticket=' + ticket;
    writeAudit_(session, 'ISSUE_RECEIPT', 'Record', rid, receiptNo);
    return { ok: true, data: { receiptNo: receiptNo, url: url, ticket: ticket } };
  });
}

/**
 * ขั้นตอนชำระเงิน — the last step of a treatment record. payload:
 * {status: 'paid'|'unpaid'|'cancelled', paymentMethod}. Moving into or out of
 * ยกเลิก changes the doctor's daily-minimum group, so that is recomputed.
 */
function payRecord(token, rid, payload) {
  return safeCall_(function () {
    var session = requireAuth_(token, 'staff');
    payload = payload || {};
    var status = payload.status || 'paid';
    if (['paid', 'unpaid', 'cancelled'].indexOf(status) === -1) throw ValidationError_('สถานะการชำระเงินไม่ถูกต้อง');
    if (status === 'paid') requireFields_(payload, { paymentMethod: 'วิธีชำระเงิน' });
    var lock = LockService.getScriptLock();
    lock.waitLock(30000);
    try {
      var sheet = getSheet_(CONFIG.SHEETS.RECORDS);
      var record = getRowById_(sheet, CONFIG.HEADERS.Records, rid);
      if (!record) throw ValidationError_('ไม่พบข้อมูลการรักษา');
      var oldStatus = record.Status || 'unpaid';
      var updated = Object.assign({}, record, { Status: status, PaymentMethod: payload.paymentMethod || record.PaymentMethod || '', UpdatedAt: nowStr_() });
      // Paid = เสร็จสิ้น; taking a payment back puts it at รอชำระเงิน again.
      if (status === 'paid') updated.VisitStatus = 'completed';
      else if (status === 'unpaid' && String(record.VisitStatus || '') === 'completed') updated.VisitStatus = 'waiting_payment';
      updateRowByIndex_(sheet, CONFIG.HEADERS.Records, record._row, updated);
      if ((oldStatus === 'cancelled') !== (status === 'cancelled')) recomputeRecordDoctors_(null, updated);
      writeAudit_(session, 'PAY', 'Record', rid, status + ' ' + (updated.PaymentMethod || ''));
      bumpDashboardCacheVersion_();
      return { ok: true, data: mapRecordOut_(updated) };
    } finally {
      lock.releaseLock();
    }
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
  var useCanula = isTrue_(r.UseCanula);
  var canulaAmount = Number(r.CanulaAmount) || 0;
  var otherAmount = Number(r.OtherAmount) || 0;
  var amount = Number(r.Amount) || 0;
  var promoAmount = Number(r.PromoAmount) || 0;
  var medicineAmount = Number(r.MedicineAmount) || 0;
  var patient = getRows_(getSheet_(CONFIG.SHEETS.PATIENTS)).filter(function (p) { return p.ID === r.PatientID; })[0];
  return {
    receiptNo: r.ReceiptNo, date: dateStr_(r.Date), patientId: r.PatientID || '', hn: patient ? (patient.HN || '') : '', patientName: r.PatientName, phone: r.Phone,
    procedureItems: parseLineItems_(r.ProcedureJSON),
    serviceCategory: r.ServiceCategory, detail: r.Detail, doctorName: r.DoctorName, paymentMethod: r.PaymentMethod,
    note: r.Note || '',
    supplies: supplies, useCanula: useCanula, canulaAmount: canulaAmount, otherAmount: otherAmount,
    // Round 13 itemized groups — see Records.PromoJSON/MedicineJSON/
    // OtherItemsJSON schema comment. Printed as their own receipt lines
    // alongside the existing supplies/canula lines.
    promoItems: parseLineItems_(r.PromoJSON), promoAmount: promoAmount,
    medicineItems: splitMedicineLines_(r).charged.map(stripCost_), medicineAmount: medicineAmount, // only the charged ones print
    oralMedItems: parseLineItems_(r.OralMedJSON), oralMedAmount: Number(r.OralMedAmount) || 0,
    canulaItems: canulaItemsOf_(r),
    otherItems: parseLineItems_(r.OtherItemsJSON),
    // `amount` is จำนวนเงินหัตถการ (base procedure price, no longer includes
    // canula — see the revenue-split engine); `grandTotal` is what actually
    // gets printed as the amount due, matching totalCollected in mapRecordOut_.
    amount: amount, grandTotal: round2_(amount + promoAmount + (useCanula ? canulaAmount : 0) + medicineAmount + (Number(r.OralMedAmount) || 0) + otherAmount),
    clinic: clinicInfo
  };
}

/* ======================================================================
   บันทึกค่าใช้จ่ายคลีนิค — Expenses (tab: ค่าใช้จ่าย) + Orders/OrderLines (tab: จัดซื้อ)
   ====================================================================== */

/* ======================================================================
   จัดยา-เบิกยา — MedWithdrawals. Medicine taken out of clinic stock for a
   patient by a sales person (step 2 of the flow). Each line then waits for
   the patient's บันทึกการรักษา to charge it (RecordID). ราคาทุน is admin-only.
   ====================================================================== */

/** A JSON array of Drive file ids (รูปยาที่เบิก) -> a clean array. */
function parseIdList_(json) {
  try {
    var list = JSON.parse(json || '[]');
    return Array.isArray(list) ? cleanPhotoIds_(list) : [];
  } catch (e) { return []; }
}
/** Keeps well-formed, distinct Drive file ids only (at most 8). */
function cleanPhotoIds_(list) {
  var seen = {}, out = [];
  (Array.isArray(list) ? list : []).forEach(function (id) {
    id = String(id || '');
    if (/^[A-Za-z0-9_-]{10,}$/.test(id) && !seen[id] && out.length < 8) { seen[id] = true; out.push(id); }
  });
  return out;
}

/** `showCost` must only be true for admins. */
function mapWithdrawalOut_(w, showCost, stockById) {
  var out = {
    id: w.WithdrawID, date: dateStr_(w.Date), patientId: w.PatientID, patientName: w.PatientName, phone: w.Phone,
    salesId: w.SalesID, salesName: w.SalesName, itemId: w.ItemID, itemName: w.ItemName, qty: Number(w.Qty) || 0,
    recordId: w.RecordID || '', note: w.Note || '', createdBy: w.CreatedBy, photoIds: parseIdList_(w.PhotoFileIds)
  };
  if (showCost) { out.unitCost = Number(w.UnitCost) || 0; out.amount = Number(w.Amount) || 0; }
  return out;
}

/**
 * Points the withdrawals record `rid` took in (its WithdrawJSON lines) at it
 * (RecordID = rid); any that pointed at rid but no longer appear go back to
 * waiting (RecordID blank). Their ราคาขาย stays what was set at จัดยา-เบิกยา —
 * the record doesn't charge them. Must run inside the script lock.
 */
function linkWithdrawals_(sheet, rows, rid, withdrawItems) {
  var taken = {};
  (withdrawItems || []).forEach(function (l) { taken[l.withdrawId] = true; });
  rows.forEach(function (w) {
    var want = taken[w.WithdrawID] ? rid : (w.RecordID === rid ? '' : w.RecordID);
    if (want !== (w.RecordID || '')) updateRowByIndex_(sheet, CONFIG.HEADERS.MedWithdrawals, w._row, Object.assign({}, w, { RecordID: want }));
  });
}

/** The patient's จัดยา-เบิกยา lines no record has charged yet — the record modal pre-fills its รายการยา from these. */
function listPendingWithdrawals(token, patientId) {
  return safeCall_(function () {
    requireAuth_(token, 'staff');
    var stockById = {};
    getRows_(getSheet_(CONFIG.SHEETS.STOCK)).forEach(function (s) { stockById[s.ItemID] = s; });
    var rows = getRows_(getSheet_(CONFIG.SHEETS.MED_WITHDRAWALS)).filter(function (w) { return w.PatientID === patientId && !w.RecordID; });
    rows.sort(function (a, b) { return String(a.CreatedAt).localeCompare(String(b.CreatedAt)); });
    return { ok: true, data: rows.map(function (w) { return mapWithdrawalOut_(w, false, stockById); }) };
  });
}

function listMedWithdrawals(token, opts) {
  return safeCall_(function () {
    var session = requireAuth_(token, 'staff');
    opts = opts || {};
    var filters = compactFilters_(opts.filters);
    var rows = getRows_(getSheet_(CONFIG.SHEETS.MED_WITHDRAWALS));
    var hnById = patientHnById_();
    if (filters.q) {
      var q = String(filters.q).trim().toLowerCase(), qDigits = q.replace(/[^\d]/g, '');
      rows = rows.filter(function (r) {
        return String(r.PatientName).toLowerCase().indexOf(q) > -1 || (qDigits && String(r.Phone).replace(/[^\d]/g, '').indexOf(qDigits) > -1) ||
          String(hnById[r.PatientID] || '').toLowerCase().indexOf(q) > -1;
      });
    }
    if (filters.salesId) rows = rows.filter(function (r) { return r.SalesID === filters.salesId; });
    var showCost = session.level === 'admin';
    var stockById = {};
    getRows_(getSheet_(CONFIG.SHEETS.STOCK)).forEach(function (s) { stockById[s.ItemID] = s; });
    // One entry per save (BatchID; rows from before it share date + patient +
    // sales): the list shows the patient/sales/status once with a row per ยา,
    // and pages by these groups so a save never splits across pages.
    var groups = {}, order = [];
    var recStatus = {};
    getRows_(getSheet_(CONFIG.SHEETS.RECORDS)).forEach(function (rr) { recStatus[rr.RID] = visitStatusOf_(rr); });
    rows.sort(function (a, b) { return String(a.CreatedAt).localeCompare(String(b.CreatedAt)) || String(a.WithdrawID).localeCompare(String(b.WithdrawID)); });
    rows.forEach(function (r) {
      var key = r.BatchID || (dateStr_(r.Date) + '|' + r.PatientID + '|' + r.SalesID);
      var g = groups[key];
      if (!g) {
        g = groups[key] = { id: key, date: dateStr_(r.Date), patientId: r.PatientID, hn: hnById[r.PatientID] || '', patientName: r.PatientName, phone: r.Phone, salesId: r.SalesID, salesName: r.SalesName, note: r.Note || '', createdAt: String(r.CreatedAt), lines: [] };
        order.push(g);
      }
      var out = mapWithdrawalOut_(r, showCost, stockById);
      out.recordStatus = recStatus[r.RecordID] || 'dispensed'; // same words as the record's สถานะการรักษา; no record yet = จัดยาแล้ว
      g.lines.push(out);
      g.createdAt = String(r.CreatedAt);
      if (!g.note && r.Note) g.note = r.Note;
      if (!g.photoIds && out.photoIds.length) g.photoIds = out.photoIds; // รูปยาที่เบิก, kept on every line of the save
    });
    order.forEach(function (g) { g.photoIds = g.photoIds || []; });
    // สถานะ (filter): the step of the group's record(s) — the earliest one if its lines sit on different records.
    order.forEach(function (g) {
      g.visitStatus = g.lines.reduce(function (min, l) { return VISIT_FLOW_.indexOf(l.recordStatus) < VISIT_FLOW_.indexOf(min) ? l.recordStatus : min; }, 'completed');
    });
    if (filters.visitStatus) order = order.filter(function (g) { return g.visitStatus === filters.visitStatus; });
    order.sort(function (a, b) { return b.date.localeCompare(a.date) || b.createdAt.localeCompare(a.createdAt); });
    return { ok: true, data: paginate_(order, opts.page, opts.pageSize) };
  });
}

/**
 * Moves a treatment record to another step: to = 'dispensed' | 'waiting_doctor' | 'with_doctor' | 'waiting_payment'.
 * Arriving (รอพบแพทย์) takes today's next queue number, and a record dispensed on an earlier day takes today as its
 * treatment date; พบแพทย์ needs the doctor (doctorId) the customer goes in to see.
 */
function setVisitStatus(token, rid, to, doctorId) {
  return safeCall_(function () {
    var session = requireAuth_(token, 'staff');
    var lock = LockService.getScriptLock();
    lock.waitLock(30000);
    try {
      var sheet = getSheet_(CONFIG.SHEETS.RECORDS);
      var row = getRowById_(sheet, CONFIG.HEADERS.Records, rid);
      if (!row) throw ValidationError_('ไม่พบข้อมูลการรักษา');
      var obj = Object.assign({}, row);
      var doctor = null;
      if (doctorId) {
        var d = getRows_(getSheet_(CONFIG.SHEETS.DOCTORS)).filter(function (x) { return x.DoctorID === doctorId; })[0];
        if (!d) throw ValidationError_('ไม่พบข้อมูลแพทย์ที่เลือก');
        doctor = { id: d.DoctorID, name: d.Name };
      }
      applyVisitStatus_(obj, to, todayStr_(), true, doctor);
      obj.UpdatedAt = nowStr_();
      updateRowByIndex_(sheet, CONFIG.HEADERS.Records, row._row, obj);
      if (dateStr_(row.Date) !== dateStr_(obj.Date)) recomputeRecordDoctors_(row, obj);
      writeAudit_(session, 'STATUS', 'Record', rid, visitStatusOf_(row) + ' → ' + to);
      bumpDashboardCacheVersion_();
      return { ok: true, data: mapRecordOut_(obj) };
    } finally {
      lock.releaseLock();
    }
  });
}

/**
 * The customers waiting to pay right now ({id, queue, patientName}) — polled by every open screen so the board's menu can
 * flag a customer who has just moved from พบแพทย์ to รอชำระเงิน.
 */
function getQueueAlerts(token) {
  return safeCall_(function () {
    requireAuth_(token, 'staff');
    var list = [];
    getRows_(getSheet_(CONFIG.SHEETS.RECORDS)).forEach(function (r) {
      if (r.Status === 'cancelled' || visitStatusOf_(r) !== 'waiting_payment') return;
      list.push({ id: r.RID, queue: Number(r.QueueDoctor) || null, patientName: r.PatientName });
    });
    return { ok: true, data: { waitingPayment: list } };
  });
}

/**
 * The "ระหว่างดำเนินการ" board for one day (default today): five columns by สถานะการรักษา with counts.
 * จัดยาแล้ว lists everything still waiting for its customer (any date — dispensing is often the day before);
 * the other columns hold that day's queue (รอพบแพทย์ / พบแพทย์ / รอชำระเงิน carried over from earlier days show on today's board, flagged).
 */
function getQueueBoard(token, date) {
  return safeCall_(function () {
    requireAuth_(token, 'staff');
    var today = todayStr_(), day = date || today;
    requireDateStr_(day, 'วันที่');
    var hnById = patientHnById_();
    var cols = { dispensed: [], waiting_doctor: [], with_doctor: [], waiting_payment: [], completed: [] };
    getRows_(getSheet_(CONFIG.SHEETS.RECORDS)).forEach(function (r) {
      if (r.Status === 'cancelled') return;
      var st = visitStatusOf_(r), d = dateStr_(r.QueueDate || r.Date), carry = false;
      if (st !== 'dispensed' && d !== day) {
        if (st !== 'completed' && day === today && d < today) carry = true; else return;
      }
      var wd = parseLineItems_(r.WithdrawJSON);
      var docs = {};
      recordDoctorShares_(r).forEach(function (x) { if (x.name) docs[x.name] = true; });
      cols[st].push({
        id: r.RID, patientId: r.PatientID, patientName: r.PatientName, hn: hnById[r.PatientID] || '', phone: r.Phone,
        date: dateStr_(r.Date), queueDate: r.QueueDate ? dateStr_(r.QueueDate) : '', carryover: carry,
        queueDoctor: Number(r.QueueDoctor) || null,
        visitDoctorId: r.VisitDoctorID || '', visitDoctorName: r.VisitDoctorName || '',
        salesName: r.SalesName || '', doctors: st === 'dispensed' || st === 'waiting_doctor' ? [] : Object.keys(docs), // no treatment record before พบแพทย์
        items: wd.slice(0, 3).map(function (l) { return l.name + ' ×' + l.qtyDispensed; }), moreItems: Math.max(0, wd.length - 3),
        total: st === 'waiting_payment' || st === 'completed' ? recordCollected_(r) : null,
        paymentMethod: r.PaymentMethod || '', createdAt: String(r.CreatedAt)
      });
    });
    var num = function (v) { return v == null ? 1e9 : v; };
    cols.dispensed.sort(function (a, b) { return a.date.localeCompare(b.date) || a.createdAt.localeCompare(b.createdAt); });
    ['waiting_doctor', 'with_doctor'].forEach(function (k) {
      cols[k].sort(function (a, b) { return a.queueDate.localeCompare(b.queueDate) || num(a.queueDoctor) - num(b.queueDoctor) || a.createdAt.localeCompare(b.createdAt); });
    });
    cols.waiting_payment.sort(function (a, b) { return a.queueDate.localeCompare(b.queueDate) || num(a.queueDoctor) - num(b.queueDoctor) || a.createdAt.localeCompare(b.createdAt); });
    cols.completed.sort(function (a, b) { return num(a.queueDoctor) - num(b.queueDoctor) || a.createdAt.localeCompare(b.createdAt); });
    return { ok: true, data: { date: day, columns: cols, counts: { dispensed: cols.dispensed.length, waiting_doctor: cols.waiting_doctor.length, with_doctor: cols.with_doctor.length, waiting_payment: cols.waiting_payment.length, completed: cols.completed.length } } };
  });
}

/** payload: {date, patientId | patientName, salesId, note, photoFileIds?, lines:[{itemId, qty}]} (no patientId: the typed patientName becomes a new patient "name (รออัพเดท)") — one MedWithdrawals row per line (all sharing a BatchID), each taken out of Stock, plus the บันทึกการรักษา (จัดยาแล้ว) they belong to. Returns {ids, recordId}. */
function saveMedWithdrawal(token, payload) {
  return safeCall_(function () {
    var session = requireAuth_(token, 'staff');
    payload = payload || {};
    requireFields_(payload, { date: 'วันที่', salesId: 'ชื่อเซลล์' });
    if (!payload.patientId && !String(payload.patientName || '').trim()) throw ValidationError_('กรุณาเลือกคนไข้หรือพิมพ์ชื่อคนไข้');
    requireDateStr_(payload.date, 'วันที่');
    var lines = (Array.isArray(payload.lines) ? payload.lines : []).filter(function (l) { return l && l.itemId; });
    if (!lines.length) throw ValidationError_('กรุณาเลือกยาอย่างน้อย 1 รายการ');

    var lock = LockService.getScriptLock();
    lock.waitLock(30000);
    beginWriteScope_();
    try {
      var patient = payload.patientId ? getRows_(getSheet_(CONFIG.SHEETS.PATIENTS)).filter(function (p) { return p.ID === payload.patientId; })[0] : null;
      if (payload.patientId && !patient) throw ValidationError_('ไม่พบข้อมูลคนไข้ที่เลือก');
      var sales = getRows_(getSheet_(CONFIG.SHEETS.SALES)).filter(function (s) { return s.SalesID === payload.salesId; })[0];
      if (!sales) throw ValidationError_('ไม่พบข้อมูลเซลล์ที่เลือก');
      var stockById = {};
      getRows_(getSheet_(CONFIG.SHEETS.STOCK)).forEach(function (r) { stockById[r.ItemID] = r; });
      var newPatientId = '';
      if (!patient) { patient = createPendingPatient_(session, payload.patientName); newPatientId = patient.ID; } // name not in the list: opens its history, to fill in later
      var patientName = patientFullName_(patient);
      var take = {};
      var rows = lines.map(function (l, i) {
        var line = ' (บรรทัดที่ ' + (i + 1) + ')';
        var item = stockById[l.itemId];
        if (!item || stockTypeOf_(item) !== 'medicine') throw ValidationError_('กรุณาเลือกยาจากรายการสต๊อกยา' + line);
        var qty = requireNumber_(l.qty, 'จำนวน' + line, { min: 0 });
        if (!qty) throw ValidationError_('จำนวนต้องมากกว่า 0' + line);
        var cost = Number(item.CostPrice) || 0;
        take[item.ItemID] = round2_((take[item.ItemID] || 0) + qty);
        return {
          Date: payload.date, PatientID: patient.ID, PatientName: patientName, Phone: patient.Phone,
          SalesID: sales.SalesID, SalesName: sales.Name, ItemID: item.ItemID, ItemName: item.Name,
          Qty: qty, UnitCost: cost, Amount: round2_(qty * cost), Note: payload.note || '', CreatedBy: session.username, CreatedAt: nowStr_(),
          RecordID: '', PhotoFileIds: JSON.stringify(cleanPhotoIds_(payload.photoFileIds))
        };
      });
      applyStockEffects_(session, {}, take, 'จัดยา-เบิกยา · ' + patientName + ' · ' + sales.Name, 'withdraw');
      var sheet = getSheet_(CONFIG.SHEETS.MED_WITHDRAWALS);
      var ids = nextSeqBlock_('WITHDRAW', rows.length, 5);
      var batchId = 'WB' + nextSeq_('WITHDRAW_BATCH', 5); // groups this save's lines
      // Every save also opens its บันทึกการรักษา (สถานะ จัดยาแล้ว), which the
      // lines belong to; staff add หัตถการ / รายการอื่นๆ there and finish it.
      var rid = 'R' + nextSeq_('RECORD', 6);
      rows.forEach(function (r, i) { r.WithdrawID = 'WD' + ids[i]; r.BatchID = batchId; r.RecordID = rid; });
      appendRows_(sheet, CONFIG.HEADERS.MedWithdrawals, rows); // all the lines in one write
      var rec = {
        RID: rid, Date: payload.date, PatientID: patient.ID, PatientName: patientName, Phone: patient.Phone,
        SalesID: sales.SalesID, SalesName: sales.Name, Status: 'unpaid', PaymentMethod: '', ReceiptNo: '', Note: '',
        CreatedBy: session.username, CreatedAt: nowStr_(), UpdatedAt: nowStr_(),
        SalesCommissionPct: Number(sales.CommissionRate) || 0, SalesShare: 0, VisitStatus: 'dispensed', StockApplied: 'TRUE',
        WithdrawJSON: JSON.stringify(rows.map(function (r) {
          return { withdrawId: r.WithdrawID, itemId: r.ItemID, name: r.ItemName, qtyDispensed: r.Qty, qtyUsed: r.Qty, qtyDeposit: 0, expiry: '' };
        }))
      };
      applyDoctorShares_(rec, []); // no doctor yet: หัตถการ lines name theirs
      appendRow_(getSheet_(CONFIG.SHEETS.RECORDS), CONFIG.HEADERS.Records, rec);
      writeAudit_(session, 'CREATE', 'MedWithdrawal', rows.map(function (r) { return r.WithdrawID; }).join(','), patientName + ' x' + rows.length + ' → ' + rid);
      bumpDashboardCacheVersion_();
      return { ok: true, data: { ids: rows.map(function (r) { return r.WithdrawID; }), recordId: rid, newPatientId: newPatientId } };
    } finally {
      endWriteScope_();
      lock.releaseLock();
    }
  });
}

/**
 * Sets the รูปยาที่เบิก of one จัดยา-เบิกยา save: `ids` are its line ids (WithdrawID),
 * `fileIds` the Drive photos (all the lines carry the same list). Staff may do this.
 */
function setWithdrawalPhotos(token, ids, fileIds) {
  return safeCall_(function () {
    var session = requireAuth_(token, 'staff');
    var want = {};
    (Array.isArray(ids) ? ids : []).forEach(function (id) { want[String(id)] = true; });
    if (!Object.keys(want).length) throw ValidationError_('ไม่พบรายการเบิกยา');
    var photos = cleanPhotoIds_(fileIds);
    var lock = LockService.getScriptLock();
    lock.waitLock(30000);
    try {
      var sheet = getSheet_(CONFIG.SHEETS.MED_WITHDRAWALS);
      var rows = getRows_(sheet).filter(function (r) { return want[r.WithdrawID]; });
      if (!rows.length) throw ValidationError_('ไม่พบรายการเบิกยา');
      rows.forEach(function (r) { updateRowByIndex_(sheet, CONFIG.HEADERS.MedWithdrawals, r._row, Object.assign({}, r, { PhotoFileIds: JSON.stringify(photos) })); });
      writeAudit_(session, 'UPDATE', 'MedWithdrawal', rows.map(function (r) { return r.WithdrawID; }).join(','), 'รูปยา ' + photos.length + ' รูป');
      return { ok: true, data: { photoIds: photos } };
    } finally {
      lock.releaseLock();
    }
  });
}

/** True when a Records row has anything on it besides its จัดยา-เบิกยา lines (so it must be kept). */
function recordHasContent_(r) {
  var some = function (json) { return parseLineItems_(json).length > 0; };
  return some(r.SuppliesJSON) || some(r.ProcedureJSON) || some(r.PromoJSON) || some(r.MedicineJSON) || some(r.OwnMedJSON) ||
    some(r.DepositUseJSON) || some(r.MessengerJSON) || some(r.OtherItemsJSON) || some(r.OralMedJSON) || (isTrue_(r.UseCanula) && (Number(r.CanulaAmount) > 0 || some(r.CanulaJSON))) ||
    !!r.OpdImageFileId || !!String(r.Detail || '').trim() || !!String(r.Note || '').trim() || (r.Status && r.Status !== 'unpaid');
}

/**
 * Admin-only; puts the line's qty back into Stock and takes it off its
 * บันทึกการรักษา (which goes too if that was all it held). Refused once the
 * record is เสร็จสิ้น, or if part of the line was put into ยาฝาก.
 */
function deleteMedWithdrawal(token, id) {
  return safeCall_(function () {
    var session = requireAuth_(token, 'admin');
    var lock = LockService.getScriptLock();
    lock.waitLock(30000);
    try {
      var sheet = getSheet_(CONFIG.SHEETS.MED_WITHDRAWALS);
      var row = getRows_(sheet).filter(function (r) { return r.WithdrawID === id; })[0];
      if (!row) throw ValidationError_('ไม่พบรายการ');
      var recSheet = getSheet_(CONFIG.SHEETS.RECORDS);
      var rec = row.RecordID ? getRows_(recSheet).filter(function (r) { return r.RID === row.RecordID; })[0] : null;
      var recUpdate = null;
      if (rec) {
        if (visitStatusOf_(rec) === 'waiting_payment' || visitStatusOf_(rec) === 'completed') throw ValidationError_('บันทึกการรักษา ' + rec.RID + ' อยู่ในขั้นรอชำระเงิน/เสร็จสิ้นแล้ว — ลบรายการยานี้ออกจากบันทึกการรักษานั้นก่อน');
        var items = parseLineItems_(rec.WithdrawJSON);
        var mine = items.filter(function (l) { return l.withdrawId === id; })[0];
        if (mine && Number(mine.qtyDeposit) > 0) throw ValidationError_('รายการนี้มียาฝากในบันทึกการรักษา ' + rec.RID + ' — แก้ที่บันทึกการรักษาก่อน');
        recUpdate = Object.assign({}, rec, { WithdrawJSON: JSON.stringify(items.filter(function (l) { return l.withdrawId !== id; })), UpdatedAt: nowStr_() });
      }
      var give = {};
      give[row.ItemID] = Number(row.Qty) || 0;
      applyStockEffects_(session, give, {}, 'ลบรายการจัดยา-เบิกยา ' + id + ' · ' + row.PatientName, 'withdraw');
      deleteRowByIndex_(sheet, row._row);
      if (recUpdate) {
        if (recordHasContent_(recUpdate) || parseLineItems_(recUpdate.WithdrawJSON).length) updateRowByIndex_(recSheet, CONFIG.HEADERS.Records, recUpdate._row, recUpdate);
        else deleteRowByIndex_(recSheet, recUpdate._row);
        bumpDashboardCacheVersion_();
      }
      writeAudit_(session, 'DELETE', 'MedWithdrawal', id, row.PatientName + ' ' + row.ItemName);
      return { ok: true };
    } finally {
      lock.releaseLock();
    }
  });
}

/* ======================================================================
   รายงาน (admin only) — opts: {dateFrom, dateTo, salesId}; blank salesId =
   everyone. Both return flat rows, oldest first, plus totals.
   ====================================================================== */

function reportOpts_(opts) {
  opts = opts || {};
  if (opts.dateFrom) requireDateStr_(opts.dateFrom, 'วันที่เริ่มต้น');
  if (opts.dateTo) requireDateStr_(opts.dateTo, 'วันที่สิ้นสุด');
  return function (date, salesId) {
    return (!opts.salesId || salesId === opts.salesId) && (!opts.dateFrom || date >= opts.dateFrom) && (!opts.dateTo || date <= opts.dateTo);
  };
}

/** The receipt's lines for one Records row, as short "name ×qty · amount" texts (same groups Receipt.html prints). */
function receiptLines_(r) {
  var out = [];
  var procs = parseLineItems_(r.ProcedureJSON);
  if (procs.length) procs.forEach(function (p) { out.push({ name: p.name, amount: Number(p.amount) || 0 }); });
  else if (Number(r.Amount)) out.push({ name: r.Detail ? 'หัตถการ (' + r.Detail + ')' : 'หัตถการ', amount: Number(r.Amount) || 0 });
  parseLineItems_(r.PromoJSON).forEach(function (p) { out.push({ name: 'โปร ' + p.name, amount: Number(p.amount) || 0 }); });
  parseLineItems_(r.MedicineJSON).forEach(function (m) { if (Number(m.amount) > 0) out.push({ name: (m.name || 'ยา') + (m.qtyDispensed ? ' ×' + m.qtyDispensed : ''), amount: Number(m.amount) || 0 }); });
  parseLineItems_(r.OralMedJSON).forEach(function (m) { if (Number(m.amount) > 0) out.push({ name: 'ยาทาน ' + (m.name || '') + (m.qty ? ' ×' + m.qty : ''), amount: Number(m.amount) || 0 }); });
  if (isTrue_(r.UseCanula) && Number(r.CanulaAmount)) {
    var canulas = canulaItemsOf_(r);
    if (canulas.length) canulas.forEach(function (c) { out.push({ name: 'Canula (เข็มทู่)' + (c.name ? ' ' + c.name : '') + (Number(c.qty) > 1 ? ' ×' + c.qty : ''), amount: Number(c.amount) || 0 }); });
    else out.push({ name: 'Canula (เข็มทู่)', amount: Number(r.CanulaAmount) || 0 });
  }
  var others = parseLineItems_(r.OtherItemsJSON);
  if (others.length) others.forEach(function (o) { out.push({ name: o.label || 'อื่นๆ', amount: Number(o.amount) || 0 }); });
  else if (Number(r.OtherAmount)) out.push({ name: 'รายการอื่นๆ', amount: Number(r.OtherAmount) || 0 });
  return out;
}

/** รายงานรายวัน: one row per treatment record (cancelled ones left out) — ชื่อคนไข้, เซลล์, the receipt's lines and its total. */
function getDailyReport(token, opts) {
  return safeCall_(function () {
    requireAuth_(token, 'admin');
    var keep = reportOpts_(opts);
    var rows = getRows_(getSheet_(CONFIG.SHEETS.RECORDS)).filter(function (r) { return r.Status !== 'cancelled' && keep(dateStr_(r.Date), r.SalesID); });
    rows.sort(function (a, b) { return dateStr_(a.Date).localeCompare(dateStr_(b.Date)) || String(a.CreatedAt).localeCompare(String(b.CreatedAt)); });
    var out = rows.map(function (r) {
      var m = mapRecordOut_(r);
      return { id: r.RID, date: m.date, patientName: r.PatientName, phone: r.Phone, salesName: r.SalesName || '', receiptNo: r.ReceiptNo || '', status: m.status, lines: receiptLines_(r), amount: m.totalCollected };
    });
    return { ok: true, data: { rows: out, totalAmount: round2_(out.reduce(function (s, x) { return s + x.amount; }, 0)) } };
  });
}

/**
 * The รายงานรายเซลล์ lines — what the clinic collects from the sales people:
 * every จัดยา-เบิกยา line and every ยาเบิกเพิ่ม of a บันทึกการรักษา, valued at
 * ราคาทุน on the day. keep(date, salesId) picks the lines. Also feeds the
 * dashboard's รายได้จากเซลล์.
 */
function salesReportRows_(keep, src) {
    var stockById = {};
    getRows_(getSheet_(CONFIG.SHEETS.STOCK)).forEach(function (s) { stockById[s.ItemID] = s; });
    var out = [];
    var push = function (o) {
      o.unit = (stockById[o.itemId] || {}).Unit || '';
      o.costAmount = round2_(o.qty * o.unitCost); // เรียกเก็บจากเซลล์
      // รายการซื้อยาเพิ่มหน้าคลีนิค (คิดค่ายา): what the customer paid less the ราคาทุน goes back to the sales person.
      o.chargedAmount = o.chargedAmount || 0;
      o.refundAmount = o.chargedAmount > 0 ? round2_(o.chargedAmount - o.costAmount) : 0;
      out.push(o);
    };
    (src ? src.withdrawals : getRows_(getSheet_(CONFIG.SHEETS.MED_WITHDRAWALS))).forEach(function (w) {
      var d = dateStr_(w.Date);
      if (!keep(d, w.SalesID)) return;
      push({
        date: d, sortKey: String(w.CreatedAt), patientName: w.PatientName, salesId: w.SalesID, salesName: w.SalesName || '', itemId: w.ItemID, itemName: w.ItemName,
        qty: Number(w.Qty) || 0, unitCost: Number(w.UnitCost) || 0, remark: ''
      });
    });
    (src ? src.records : getRows_(getSheet_(CONFIG.SHEETS.RECORDS))).forEach(function (r) {
      var d = dateStr_(r.Date);
      if (r.Status === 'cancelled' || !isTrue_(r.UseMedicine) || !keep(d, r.SalesID)) return;
      parseLineItems_(r.MedicineJSON).forEach(function (l) {
        if (!('qtyDispensed' in l) || !l.itemId) return;
        var s = stockById[l.itemId] || {};
        push({
          date: d, sortKey: String(r.CreatedAt), patientName: r.PatientName, salesId: r.SalesID || '', salesName: r.SalesName || '', itemId: l.itemId, itemName: l.name,
          qty: Number(l.qtyDispensed) || 0, unitCost: l.unitCost != null ? Number(l.unitCost) || 0 : Number(s.CostPrice) || 0,
          chargedAmount: isChargedMedLine_(l) ? Number(l.amount) || 0 : 0,
          remark: (isChargedMedLine_(l) ? 'ซื้อเพิ่มหน้าคลีนิค (คิดค่ายา)' : 'ซื้อเพิ่มหน้าคลีนิค (ไม่คิดค่ายา)') + ' · บันทึกการรักษา ' + r.RID
        });
      });
    });
    out.sort(function (a, b) { return a.date.localeCompare(b.date) || a.sortKey.localeCompare(b.sortKey); });
    out.forEach(function (o) { delete o.sortKey; });
    return out;
}

/**
 * รายงานรายเซลล์: the ราคาทุน the clinic owes each sales person — the customer
 * pays the sales person for the ยา directly, so no selling price is shown.
 * One row per ยา: every จัดยา-เบิกยา line, plus every ยาเบิกเพิ่ม line on a
 * non-cancelled treatment record (หมายเหตุ says which record it was added
 * on). ราคาทุน is the snapshot taken when it left stock. Also totals per
 * sales person.
 */
function getSalesReport(token, opts) {
  return safeCall_(function () {
    requireAuth_(token, 'admin');
    var keep = reportOpts_(opts);
    var out = salesReportRows_(keep);
    var bySales = {}, order = [];
    out.forEach(function (o) {
      var key = o.salesId || o.salesName || '-';
      if (!bySales[key]) { bySales[key] = { salesId: o.salesId, salesName: o.salesName || '-', qty: 0, cost: 0, refund: 0 }; order.push(key); }
      bySales[key].qty = round2_(bySales[key].qty + o.qty);
      bySales[key].cost = round2_(bySales[key].cost + o.costAmount);
      bySales[key].refund = round2_(bySales[key].refund + o.refundAmount);
    });
    var sum = function (k) { return round2_(out.reduce(function (s, o) { return s + o[k]; }, 0)); };
    return { ok: true, data: { rows: out, bySales: order.map(function (k) { return bySales[k]; }), totalQty: sum('qty'), totalCost: sum('costAmount'), totalRefund: sum('refundAmount') } };
  });
}

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
    id: o.PoID, date: dateStr_(o.Date), supplier: o.Supplier, orderType: o.OrderType === 'sales' ? 'sales' : 'other', totalAmount: Number(o.TotalAmount) || 0,
    createdBy: o.CreatedBy, row: o._row,
    lines: (lines || []).map(function (l) {
      return { id: l.LineID, category: l.Category, itemName: l.ItemName, qty: Number(l.Qty) || 0, unitPrice: Number(l.UnitPrice) || 0, lineTotal: Number(l.LineTotal) || 0, itemId: l.ItemID || '' };
    }),
    stockApplied: isTrue_(o.StockApplied)
  };
}

/**
 * Stock a PO puts in, in applyStockEffects_'s shape (negative = added to
 * stock): {itemId: -qty}. Only POs with StockApplied=TRUE count, and items
 * no longer in Stock are skipped so a deleted stock item never blocks
 * editing/deleting the PO.
 */
function orderEffects_(orderRow, lines, stockIds) {
  var eff = {};
  if (!orderRow || !isTrue_(orderRow.StockApplied)) return eff;
  lines.forEach(function (l) {
    var id = l.ItemID || l.itemId, qty = Number(l.Qty != null ? l.Qty : l.qty) || 0;
    if (id && qty && stockIds[id]) eff[id] = round2_((eff[id] || 0) - qty);
  });
  return eff;
}
function stockIdSet_() {
  var ids = {};
  getRows_(getSheet_(CONFIG.SHEETS.STOCK)).forEach(function (s) { ids[s.ItemID] = true; });
  return ids;
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

/** order: {id?, date, orderType ('sales' ยา(จากเซลล์): supplier = the sales person's name | 'other' อื่นๆ: the vendor), supplier, lines:[{category,itemName,qty,unitPrice,itemId?}]} — TotalAmount computed server-side; lines with an itemId add their qty to Stock (see orderEffects_). */
function saveOrder(token, order) {
  return safeCall_(function () {
    var session = requireAuth_(token, 'staff');
    var orderType = order.orderType === 'sales' ? 'sales' : 'other';
    requireFields_(order, { date: 'วันที่', supplier: orderType === 'sales' ? 'ชื่อเซลล์' : 'ผู้จำหน่าย' });
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
      var existingOrder = isNew ? null : getRows_(orderSheet).filter(function (o) { return o.PoID === order.id; })[0];
      if (!isNew && !existingOrder) throw ValidationError_('ไม่พบใบสั่งซื้อ');
      var poId = isNew ? ('PO' + nextSeq_('ORDER', 5)) : order.id;
      var obj = {
        PoID: poId, Date: order.date, Supplier: order.supplier, OrderType: orderType, TotalAmount: totalAmount,
        CreatedBy: isNew ? session.username : (order.createdBy || session.username),
        CreatedAt: isNew ? nowStr_() : (order.createdAt || nowStr_()),
        StockApplied: isNew ? 'TRUE' : (existingOrder.StockApplied || '')
      };
      // Add the purchased qty to Stock (an edit moves only the difference).
      var stockIds = stockIdSet_();
      var oldLines = isNew ? [] : getRows_(lineSheet).filter(function (l) { return l.PoID === order.id; });
      applyStockEffects_(session, orderEffects_(existingOrder, oldLines, stockIds), orderEffects_(obj, lines, stockIds),
        'จัดซื้อ ' + poId + ' · ' + order.supplier, 'purchase');
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
          Qty: Number(l.qty), UnitPrice: Number(l.unitPrice), LineTotal: Number(l.qty) * Number(l.unitPrice),
          ItemID: l.itemId || ''
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
    var lock = LockService.getScriptLock();
    lock.waitLock(30000);
    try {
      var orderSheet = getSheet_(CONFIG.SHEETS.ORDERS);
      var orderRow = getRows_(orderSheet).filter(function (o) { return o.PoID === id; })[0];
      if (!orderRow) throw ValidationError_('ไม่พบใบสั่งซื้อ');
      var lineSheet = getSheet_(CONFIG.SHEETS.ORDER_LINES);
      var lineRows = getRows_(lineSheet).filter(function (l) { return l.PoID === id; });
      // Take the purchased qty back out of Stock (may go negative, like any other stock-out).
      applyStockEffects_(session, orderEffects_(orderRow, lineRows, stockIdSet_()), {},
        'ลบใบสั่งซื้อ ' + id + ' · ' + orderRow.Supplier, 'purchase');
      deleteRowByIndex_(orderSheet, orderRow._row);
      lineRows.sort(function (a, b) { return b._row - a._row; });
      lineRows.forEach(function (l) { deleteRowByIndex_(lineSheet, l._row); });
      writeAudit_(session, 'DELETE', 'Order', id, '');
      bumpDashboardCacheVersion_(); // Round 19 item 4: purchase orders feed dashboard expense totals
      return { ok: true };
    } finally {
      lock.releaseLock();
    }
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
    requireAuth_(token, 'admin'); // รายได้จากเซลล์ is built from ราคาทุน, and the screen is admin-only
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

    var arch = archiveSummary_(dateFrom, dateTo); // days moved to a backup file (see backupData) still count here
    var allRecords = getRows_(getSheet_(CONFIG.SHEETS.RECORDS));
    var paidRecords = allRecords.filter(function (r) { return r.Status === 'paid'; });
    var periodRecords = paidRecords.filter(function (r) { var d = dateStr_(r.Date); return d >= dateFrom && d <= dateTo; });

    var allExpenses = getRows_(getSheet_(CONFIG.SHEETS.EXPENSES));
    var periodExpenses = allExpenses.filter(function (r) { var d = dateStr_(r.Date); return d >= dateFrom && d <= dateTo; });
    var allOrders = getRows_(getSheet_(CONFIG.SHEETS.ORDERS));
    var periodOrders = allOrders.filter(function (r) { var d = dateStr_(r.Date); return d >= dateFrom && d <= dateTo; });

    // True total money collected per record (recordCollected_) ties out
    // exactly with clinic + sales + doctor shares (recordClinicShare_).
    var recordTotalCollected_ = recordCollected_;

    // รายได้ = collected from customers (หัตถการ / โปร / Canula / อื่นๆ, paid
    // records) + collected from the sales people (รายงานรายเซลล์).
    // รายจ่าย = ค่าใช้จ่าย + จัดซื้อ + ส่วนแบ่งแพทย์ + ส่วนแบ่งเซลล์.
    var salesRows = salesReportRows_(function () { return true; });
    var periodSalesRows = salesRows.filter(function (o) { return o.date >= dateFrom && o.date <= dateTo; });
    var customerRevenue = periodRecords.reduce(function (s, r) { return s + recordTotalCollected_(r); }, 0) + arch.m.customerRevenue;
    var salesIncome = periodSalesRows.reduce(function (s, o) { return s + o.costAmount; }, 0) + arch.m.salesIncome;
    var totalRevenue = round2_(customerRevenue + salesIncome);
    var expenseFromExpenses = periodExpenses.reduce(function (s, r) { return s + (Number(r.Amount) || 0); }, 0) + arch.expenseTotal;
    var expenseFromOrders = periodOrders.reduce(function (s, r) { return s + (Number(r.TotalAmount) || 0); }, 0) + arch.m.orders;
    var totalDoctorShare = periodRecords.reduce(function (s, r) { return s + (Number(r.DoctorShare) || 0); }, 0) + arch.m.doctorShareTotal;
    var salesShareTotal = periodRecords.reduce(function (s, r) { return s + (Number(r.SalesShare) || 0); }, 0) + arch.m.salesShareTotal;
    var totalExpense = round2_(expenseFromExpenses + expenseFromOrders + totalDoctorShare + salesShareTotal);
    var profit = round2_(totalRevenue - totalExpense);
    var profitPct = totalRevenue > 0 ? (profit / totalRevenue) * 100 : 0;


    var uniquePatientIds = {};
    periodRecords.forEach(function (r) { uniquePatientIds[r.PatientID] = true; });
    Object.keys(arch.patients).forEach(function (id) { uniquePatientIds[id] = true; });
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
    salesRows.forEach(function (o) { var m = monthKey_(o.date); if (monthlyRevenue[m] !== undefined) monthlyRevenue[m] += o.costAmount; });
    paidRecords.forEach(function (r) { var m = monthKey_(dateStr_(r.Date)); if (monthlyExpense[m] !== undefined) monthlyExpense[m] += (Number(r.DoctorShare) || 0) + (Number(r.SalesShare) || 0); });
    allExpenses.forEach(function (r) { var m = monthKey_(dateStr_(r.Date)); if (monthlyExpense[m] !== undefined) monthlyExpense[m] += Number(r.Amount) || 0; });
    allOrders.forEach(function (r) { var m = monthKey_(dateStr_(r.Date)); if (monthlyExpense[m] !== undefined) monthlyExpense[m] += Number(r.TotalAmount) || 0; });
    var archMonthly = archiveMonthly_();
    var monthlyTrend = months.map(function (m) {
      var rev = monthlyRevenue[m] + ((archMonthly[m] || {}).revenue || 0), exp = monthlyExpense[m] + ((archMonthly[m] || {}).expense || 0);
      // Round 15 item 8.2: "profit" (บาท) is new, additive — the trend chart
      // now plots the profit LINE against the same money axis as the
      // revenue/expense BARS instead of a separate %-based right axis, so it
      // needs the raw money amount, not just profitPct.
      return { month: m, revenue: rev, expense: exp, profit: rev - exp, profitPct: rev > 0 ? ((rev - exp) / rev) * 100 : 0 };
    });

    // --- Revenue by type: what the clinic collects from customers
    // (หัตถการ / โปรโมชัน / Canula / อื่นๆ) and from the sales people — the
    // slices add up to รายได้รวม (อื่นๆ is the remainder, so they always do). ---
    var procRevenue = periodRecords.reduce(function (s, r) { return s + (Number(r.Amount) || 0); }, 0) + arch.m.proc;
    var promoRevenue = periodRecords.reduce(function (s, r) { return s + (Number(r.PromoAmount) || 0); }, 0) + arch.m.promo;
    var canulaRevenue = periodRecords.reduce(function (s, r) { return s + (isTrue_(r.UseCanula) ? (Number(r.CanulaAmount) || 0) : 0); }, 0) + arch.m.canula;
    var drugRevenue = periodRecords.reduce(function (s, r) { return s + (Number(r.MedicineAmount) || 0) + (Number(r.OralMedAmount) || 0); }, 0) + arch.m.drug;
    var byRevenueType = {
      'ค่าหัตถการ': procRevenue,
      'โปรโมชัน': promoRevenue,
      'Canula': canulaRevenue,
      'ค่ายา (ซื้อเพิ่ม/ยาทาน)': round2_(drugRevenue),
      'รายการอื่นๆ': round2_(customerRevenue - procRevenue - promoRevenue - canulaRevenue - drugRevenue),
      'รายได้จากเซลล์ (ยา)': round2_(salesIncome)
    };
    Object.keys(byRevenueType).forEach(function (k) { if (!(byRevenueType[k] > 0)) delete byRevenueType[k]; });

    // --- Revenue by doctor (their actual commission share, not the full record amount) ---
    var byDoctor = {};
    periodRecords.forEach(function (r) { recordDoctorShares_(r).forEach(function (d) { byDoctor[d.name] = (byDoctor[d.name] || 0) + (Number(d.share) || 0); }); });
    Object.keys(arch.doctor).forEach(function (k) { byDoctor[k] = (byDoctor[k] || 0) + arch.doctor[k]; });

    // --- Revenue by sales (their actual commission share) ---
    var bySales = {};
    periodRecords.filter(function (r) { return r.SalesID; }).forEach(function (r) { bySales[r.SalesName] = (bySales[r.SalesName] || 0) + (Number(r.SalesShare) || 0); });
    Object.keys(arch.sales).forEach(function (k) { bySales[k] = (bySales[k] || 0) + arch.sales[k]; });

    // --- Expense by category (Expenses sheet categories + a single lump "จัดซื้อ" bucket for Orders) ---
    var byExpenseCat = {};
    periodExpenses.forEach(function (r) { byExpenseCat[r.Category] = (byExpenseCat[r.Category] || 0) + (Number(r.Amount) || 0); });
    Object.keys(arch.expense).forEach(function (k) { byExpenseCat[k] = (byExpenseCat[k] || 0) + arch.expense[k]; });
    if (expenseFromOrders > 0) byExpenseCat['จัดซื้อ'] = (byExpenseCat['จัดซื้อ'] || 0) + expenseFromOrders;
    if (totalDoctorShare > 0) byExpenseCat['ส่วนแบ่งแพทย์'] = (byExpenseCat['ส่วนแบ่งแพทย์'] || 0) + totalDoctorShare;
    if (salesShareTotal > 0) byExpenseCat['ส่วนแบ่งเซลล์'] = (byExpenseCat['ส่วนแบ่งเซลล์'] || 0) + salesShareTotal;

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

    Object.keys(arch.promo).forEach(function (k) { byPromotion[k] = (byPromotion[k] || 0) + arch.promo[k]; });

    function toChartArray(map) {
      return Object.keys(map).map(function (k) { return { label: k, value: map[k] }; }).sort(function (a, b) { return b.value - a.value; });
    }

    var result = {
      period: { from: dateFrom, to: dateTo },
      stats: {
        totalRevenue: totalRevenue, totalExpense: totalExpense, profit: profit, profitPct: profitPct,
        customerCount: customerCount,
        customerRevenue: round2_(customerRevenue), salesIncome: round2_(salesIncome), // the two parts of totalRevenue
        doctorShare: round2_(totalDoctorShare), salesShare: round2_(salesShareTotal)   // expenses, with ค่าใช้จ่าย + จัดซื้อ
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
    } else if (opts.kind === 'withdraw-photo') {
      // รูปยาที่เบิก (จัดยา-เบิกยา) — same Year/Month/Day folders as the OPD photos, named after the withdrawal's date.
      var wdDate = opts.date ? new Date(opts.date) : new Date();
      if (isNaN(wdDate.getTime())) wdDate = new Date();
      folder = getOpdPhotoDayFolder_(wdDate);
      fileName = 'med_' + Utilities.formatDate(wdDate, Session.getScriptTimeZone() || 'Asia/Bangkok', 'yyyyMMdd') + '_' + (opts.patientName || 'patient') + '_' + new Date().getTime() + guessExtension_(opts.mimeType);
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
   DEMO DATA — clearAllData() empties the business data; seedDemoData()
   clears and then fills every sheet with sample data in the current format
   (stock types/prices, StockLog history, treatment records with เบิกยา/ฝาก/
   ใช้ยาฝาก lines, MedDeposits, เบิกยาคลีนิค, purchase orders linked to stock).
   Both are admin-only from the client, or run them from the Apps Script
   editor with no arguments. Never touched: the Staff sheet (login
   accounts), its STAFF counter, and ClinicSettings (company info + logo).
   ====================================================================== */

/** Empties every data row (values and formats — so no stale date formats linger); the header row is kept. */
function clearSheetData_(sheet) {
  var lastRow = sheet.getLastRow();
  if (lastRow > 1) sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn()).clear();
}

/** Every sheet except Staff and ClinicSettings, plus all counters except STAFF. */
function clearBusinessData_() {
  var keep = {};
  keep[CONFIG.SHEETS.STAFF] = true;
  keep[CONFIG.SHEETS.CLINIC_SETTINGS] = true;
  keep[CONFIG.SHEETS.COUNTERS] = true;
  Object.keys(CONFIG.SHEETS).forEach(function (key) {
    var name = CONFIG.SHEETS[key];
    if (!keep[name]) clearSheetData_(getSheet_(name));
  });
  // Counters: keep STAFF so new staff ids never collide with existing accounts.
  var counters = getSheet_(CONFIG.SHEETS.COUNTERS);
  var staffSeq = getRows_(counters).filter(function (r) { return r.Key === 'STAFF'; })[0];
  clearSheetData_(counters);
  if (staffSeq) appendRow_(counters, CONFIG.HEADERS.Counters, { Key: 'STAFF', LastSeq: staffSeq.LastSeq });
  invalidateMasterCache_();
  invalidatePatientsLiteCache_();
  invalidateStockLiteCache_();
  bumpDashboardCacheVersion_();
}

function clearAllData(token) {
  return safeCall_(function () {
    if (token) requireAuth_(token, 'admin');
    var lock = LockService.getScriptLock();
    lock.waitLock(30000);
    try {
      clearBusinessData_();
      return { ok: true, message: 'ล้างข้อมูลเรียบร้อยแล้ว (เก็บบัญชีผู้ใช้และข้อมูลบริษัทไว้)' };
    } finally {
      lock.releaseLock();
    }
  });
}

/** ราคาขาย per unit for `qty` units of a seeded medicine — same tier rule as the record form (largest priced tier <= qty, else the smallest). */
function seedTierPrice_(item, qty) {
  var pick = null, first = null;
  STOCK_PRICE_TIERS.forEach(function (t, i) {
    var p = item.prices[i];
    if (p == null || p === '') return;
    if (first == null) first = p;
    if (t <= qty) pick = p;
  });
  return pick != null ? pick : (first || 0);
}

function seedDemoData(token) {
  return safeCall_(function () {
    if (token) requireAuth_(token, 'admin');
    var lock = LockService.getScriptLock();
    lock.waitLock(30000);
    try {
      clearBusinessData_();
      var tz = Session.getScriptTimeZone() || 'Asia/Bangkok';
      var today = new Date();
      var dayStr = function (daysAgo) { var d = new Date(today.getFullYear(), today.getMonth(), today.getDate()); d.setDate(d.getDate() - daysAgo); return Utilities.formatDate(d, tz, 'yyyy-MM-dd'); };
      var actor = 'admin';

      // ---- Categories ----
      var catSheet = getSheet_(CONFIG.SHEETS.CATEGORIES);
      var cats = { treatment: ['โบท็อกซ์', 'ฟิลเลอร์', 'เลเซอร์', 'ดูแลผิวหน้า', 'ร้อยไหม'], stock: ['ยาฉีด', 'ยาทา/ยากิน', 'อุปกรณ์ทางการแพทย์', 'วัสดุสิ้นเปลือง'], expense: ['ค่าเช่า', 'ค่าน้ำค่าไฟ', 'เงินเดือนพนักงาน', 'การตลาด', 'ค่าซ่อมบำรุง', 'อื่นๆ'] };
      Object.keys(cats).forEach(function (type) {
        cats[type].forEach(function (name, i) {
          appendRow_(catSheet, CONFIG.HEADERS.Categories, { CategoryID: 'C' + nextSeq_('CATEGORY', 3), Type: type, Name: name, SortOrder: i + 1 });
        });
      });

      // ---- Doctors / Sales / Promotions ----
      var doctors = [
        { name: 'พญ.วลิน ศรีสุข', minAmount: 1500, commissionPct: 15 },
        { name: 'นพ.ธนกร รุ่งเรือง', minAmount: 1200, commissionPct: 12 },
        { name: 'พญ.ปิยะดา แสงทอง', minAmount: 1000, commissionPct: 10 }
      ].map(function (d) {
        d.id = 'D' + nextSeq_('DOCTOR', 3);
        appendRow_(getSheet_(CONFIG.SHEETS.DOCTORS), CONFIG.HEADERS.Doctors, { DoctorID: d.id, Name: d.name, Status: 'active', MinAmount: d.minAmount, CreatedAt: nowStr_(), CommissionPct: d.commissionPct });
        return d;
      });
      var sales = [
        { name: 'คุณอลิสา มั่งมี', commissionRate: 10 },
        { name: 'คุณกิตติ ขายเก่ง', commissionRate: 8 }
      ].map(function (s) {
        s.id = 'SL' + nextSeq_('SALES', 3);
        appendRow_(getSheet_(CONFIG.SHEETS.SALES), CONFIG.HEADERS.Sales, { SalesID: s.id, Name: s.name, Status: 'active', CommissionRate: s.commissionRate, MinAmount: 0, CreatedAt: nowStr_() });
        return s;
      });
      var promos = [
        { name: 'โบท็อกซ์หน้าเรียว 50 ยูนิต', price: 8900, doctorSharePct: 20, status: 'active' },
        { name: 'ฟิลเลอร์ใต้ตา 1 CC', price: 12500, doctorSharePct: 25, status: 'active' },
        { name: 'เลเซอร์กระชับผิวหน้า', price: 4500, doctorSharePct: 15, status: 'active' },
        { name: 'โปรแกรมโปรโมชันเก่า', price: 3000, doctorSharePct: 10, status: 'inactive' }
      ].map(function (p) {
        p.id = 'PM' + nextSeq_('PROMO', 3);
        appendRow_(getSheet_(CONFIG.SHEETS.PROMOTIONS), CONFIG.HEADERS.Promotions, { PromoID: p.id, Name: p.name, Price: p.price, DoctorSharePct: p.doctorSharePct, Status: p.status, CreatedAt: nowStr_(), UpdatedAt: nowStr_() });
        return p;
      });
      var activePromos = promos.filter(function (p) { return p.status === 'active'; });
      // รายการหัตถการ — one per หมวดหัตถการ (the treatment categories above).
      var procs = cats.treatment.map(function (cat, k) {
        var p = { id: 'PR' + nextSeq_('PROCEDURE', 4), name: cat + ' - โปรแกรมมาตรฐาน', category: cat, price: 1500 + k * 900 };
        appendRow_(getSheet_(CONFIG.SHEETS.PROCEDURES), CONFIG.HEADERS.Procedures, { ProcID: p.id, Name: p.name, Category: p.category, Price: p.price, Status: 'active', CreatedAt: nowStr_(), UpdatedAt: nowStr_() });
        return p;
      });

      // ---- Stock (written at the end with its final Qty) ----
      var stock = [
        { type: 'medicine', name: 'โบท็อกซ์ Dysport 500u', category: 'ยาฉีด', unit: 'ขวด', open: 30, min: 8, cost: 2500, prices: [4500, 4300, 4100, 3900, 3700, 3500] },
        { type: 'medicine', name: 'ฟิลเลอร์ Juvederm 1ml', category: 'ยาฉีด', unit: 'หลอด', open: 20, min: 5, cost: 5200, prices: [9000, 8700, 8400, 8100, '', ''] },
        { type: 'medicine', name: 'วิตามินซีฉีด 1000mg', category: 'ยาฉีด', unit: 'แอมป์', open: 60, min: 15, cost: 120, prices: [350, 330, 300, 280, 260, 240] },
        { type: 'medicine', name: 'ครีมทาแผลหลังเลเซอร์', category: 'ยาทา/ยากิน', unit: 'หลอด', open: 25, min: 6, cost: 180, prices: [450, 420, 400, '', '', ''] },
        { type: 'medicine', name: 'ยาแก้ปวด Paracetamol 500', category: 'ยาทา/ยากิน', unit: 'แผง', open: 12, min: 10, cost: 15, prices: [40, 35, 30, 28, 25, 20] },
        { type: 'supply', name: 'เข็ม Canula 25G', category: 'อุปกรณ์ทางการแพทย์', unit: 'ชิ้น', open: 150, min: 30 },
        { type: 'supply', name: 'ไหมร้อยหน้า PDO', category: 'อุปกรณ์ทางการแพทย์', unit: 'เส้น', open: 60, min: 20 },
        { type: 'supply', name: 'ถุงมือทางการแพทย์', category: 'วัสดุสิ้นเปลือง', unit: 'กล่อง', open: 40, min: 10 },
        { type: 'supply', name: 'สำลีแอลกอฮอล์', category: 'วัสดุสิ้นเปลือง', unit: 'ห่อ', open: 8, min: 10 },
        { type: 'supply', name: 'ผ้าก๊อซปลอดเชื้อ', category: 'วัสดุสิ้นเปลือง', unit: 'ห่อ', open: 0, min: 5 },
        // รายการยาทาน: one ราคาขาย each
        { type: 'oral', name: 'พาราเซตามอล 500mg (ยาทาน)', category: 'ยาทา/ยากิน', unit: 'แผง', open: 120, min: 20, prices: [35, '', '', '', '', ''] },
        { type: 'oral', name: 'อะม็อกซีซิลลิน 500mg (ยาทาน)', category: 'ยาทา/ยากิน', unit: 'แผง', open: 80, min: 15, prices: [90, '', '', '', '', ''] },
        { type: 'oral', name: 'วิตามินซี 500mg (ยาทาน)', category: 'ยาทา/ยากิน', unit: 'กระปุก', open: 40, min: 10, prices: [250, '', '', '', '', ''] }
      ];
      stock.forEach(function (s) { s.id = 'ST' + nextSeq_('STOCK', 4); s.qty = 0; });
      var meds = stock.filter(function (s) { return s.type === 'medicine'; });
      var supplies = stock.filter(function (s) { return s.type === 'supply'; });
      var orals = stock.filter(function (s) { return s.type === 'oral'; });

      // ---- Patients ----
      var patients = [
        ['สมหญิง', 'ใจดี', 'มิ้นท์', '0812345678', 'female'], ['อนันต์', 'พูลสวัสดิ์', 'บอส', '0898765432', 'male'],
        ['วิภาวดี', 'รักสวย', 'วิว', '0865551234', 'female'], ['ธีรพงษ์', 'มั่นคง', 'เต้', '0911112222', 'male'],
        ['ชนิดา', 'แสงจันทร์', 'จันทร์', '0823334444', 'female'], ['ปรีชา', 'ทองดี', 'ชา', '0955556666', 'male'],
        ['กมลชนก', 'บุญมี', 'ก้อย', '0877778888', 'female'], ['ณัฐวุฒิ', 'ศรีวิไล', 'นัท', '0899990000', 'male'],
        ['สุภาพร', 'เพชรรัตน์', 'พร', '0801231234', 'female'], ['วรากร', 'ชัยชนะ', 'กร', '0845674567', 'male']
      ].map(function (p, i) {
        var id = 'P' + nextSeq_('PATIENT', 5), hn = 'HN' + String(1001 + i);
        appendRow_(getSheet_(CONFIG.SHEETS.PATIENTS), CONFIG.HEADERS.Patients, {
          ID: id, HN: hn, FirstName: p[0], LastName: p[1], Phone: p[3], DOB: '', Gender: p[4], Email: '', PhotoFileId: '',
          JoinedDate: Utilities.formatDate(new Date(2025, i % 12, (i % 27) + 1), tz, 'yyyy-MM-dd'), CreatedAt: nowStr_(), UpdatedAt: nowStr_(), Nickname: p[2]
        });
        return { id: id, hn: hn, name: p[0] + ' ' + p[1], phone: p[3] };
      });

      // ---- Stock movements are queued as they're generated and replayed in
      //      date order at the end, so each StockLog row's QtyRemaining is the
      //      true running balance and the latest row equals Stock.Qty. Each
      //      replayed row gets a unique CreatedAt so same-day order is fixed. ----
      var logs = [], seqInDay = {}, events = [];
      var stamp = function (date) { var n = seqInDay[date] = (seqInDay[date] || 0) + 1; return date + ' 09:' + String(Math.floor(n / 60)).padStart(2, '0') + ':' + String(n % 60).padStart(2, '0'); };
      var apply = function (date, item, kind, qtyIn, qtyOut, note) {
        item.qty = round2_(item.qty + qtyIn - qtyOut);
        logs.push({ Date: date, ItemID: item.id, ItemName: item.name, QtyIn: qtyIn, QtyRemaining: item.qty, Note: note, CreatedBy: actor, CreatedAt: stamp(date), QtyOut: qtyOut, Kind: kind });
      };
      var move = function (date, item, kind, qtyIn, qtyOut, note) {
        events.push({ date: date, idx: events.length, run: function () { apply(date, item, kind, qtyIn, qtyOut, note); } });
      };
      stock.forEach(function (s) { if (s.open) move(dayStr(90), s, 'opening', s.open, 0, 'ยอดยกมา'); });

      // Purchase orders (linked to stock) — every 15 days; the last three supplies
      // are never restocked so the demo shows ใกล้หมด / หมด items.
      var orderSheet = getSheet_(CONFIG.SHEETS.ORDERS), lineSheet = getSheet_(CONFIG.SHEETS.ORDER_LINES);
      var suppliers = ['บริษัท เมดิคอล ซัพพลาย จำกัด', 'ร้านเวชภัณฑ์ไทย', 'บริษัท บิวตี้ อิมปอร์ต จำกัด'];
      for (var o = 0; o < 5; o++) {
        var poDate = dayStr(75 - o * 15), poId = 'PO' + nextSeq_('ORDER', 5), total = 0, poLines = [];
        var poType = o % 2 === 0 ? 'sales' : 'other'; // ยา(จากเซลล์): the supplier is the sales person
        var poSupplier = poType === 'sales' ? sales[(o / 2) % sales.length].name : suppliers[o % suppliers.length];
        for (var l = 0; l < 2 + (o % 2); l++) {
          var it = stock[(o * 2 + l) % 7], q = 10 + l * 5, price = it.cost || (40 + l * 25);
          total += q * price;
          poLines.push({ LineID: poId + '-' + (l + 1), PoID: poId, Category: it.category, ItemName: it.name, Qty: q, UnitPrice: price, LineTotal: q * price, ItemID: it.id });
          move(poDate, it, 'purchase', q, 0, 'จัดซื้อ ' + poId + ' · ' + poSupplier);
        }
        // One free-text line (not in stock) to show it doesn't touch stock.
        if (o === 2) { poLines.push({ LineID: poId + '-9', PoID: poId, Category: 'วัสดุสิ้นเปลือง', ItemName: 'ถุงขยะ', Qty: 5, UnitPrice: 30, LineTotal: 150, ItemID: '' }); total += 150; }
        appendRow_(orderSheet, CONFIG.HEADERS.Orders, { PoID: poId, Date: poDate, Supplier: poSupplier, OrderType: poType, TotalAmount: total, CreatedBy: actor, CreatedAt: nowStr_(), StockApplied: 'TRUE' });
        poLines.forEach(function (row) { appendRow_(lineSheet, CONFIG.HEADERS.OrderLines, row); });
      }

      // Treatment records — 16 days (oldest first), 1-4 records a day. Most start
      // as a จัดยา-เบิกยา (the withdrawal, then หัตถการ / โปร / อื่นๆ added to its
      // record), some also have ยาเบิกเพิ่ม, a few visits pay two doctors. Every visit
      // before today is finished and paid (a few cancelled); today's two first visits are
      // treated but not yet paid (รอชำระเงิน). The status pass below gives each its
      // queue number and doctor, and spreads the จัดยา-เบิกยา-only records over the steps.
      var recSheet = getSheet_(CONFIG.SHEETS.RECORDS), wdSheet = getSheet_(CONFIG.SHEETS.MED_WITHDRAWALS);
      var payMethods = ['เงินสด', 'บัตรเครดิต', 'โอนเงิน'];
      var perDay = [3, 1, 2, 4, 1, 2, 3, 1, 4, 2, 1, 3, 2, 4, 1, 3];
      var groups = {}, i = 0, seedDoctorById = {}, lots = [];
      doctors.forEach(function (d) { seedDoctorById[d.id] = { DoctorID: d.id, Name: d.name, CommissionPct: d.commissionPct }; });
      var expiryAfter = function (days) { return Utilities.formatDate(new Date(today.getFullYear(), today.getMonth(), today.getDate() + days), tz, 'yyyy-MM-dd'); };
      // ยาฝาก lots expire at different times so the รายการยาฝาก filters have something to show: the second lot already has, the third is about to.
      var lotExpiry = function () { var n = lots.length; return n === 1 ? expiryAfter(-10) : (n === 2 ? expiryAfter(12) : expiryAfter(90 + n * 30)); };
      var addWithdrawal = function (date, patient, salesP, m, q, rid, batch) {
        var wid = 'WD' + nextSeq_('WITHDRAW', 5);
        appendRow_(wdSheet, CONFIG.HEADERS.MedWithdrawals, {
          WithdrawID: wid, Date: date, PatientID: patient.id, PatientName: patient.name, Phone: patient.phone,
          SalesID: salesP.id, SalesName: salesP.name, ItemID: m.id, ItemName: m.name, Qty: q, UnitCost: m.cost, Amount: q * m.cost,
          Note: '', CreatedBy: actor, CreatedAt: nowStr_(), BatchID: batch,
          RecordID: rid
        });
        move(date, m, 'withdraw', 0, q, 'จัดยา-เบิกยา · ' + patient.name + ' · ' + salesP.name);
        return wid;
      };
      perDay.forEach(function (count, k) {
        var date = dayStr((perDay.length - 1 - k) * 4);
        for (var j = 0; j < count; j++, i++) {
          var patient = patients[i % patients.length], doctor = doctors[i % doctors.length];
          var fromWd = i % 3 !== 1; // started as a จัดยา-เบิกยา, which always has a sales person
          var salesP = (fromWd || i % 4 !== 0) ? sales[k % sales.length] : null; // one sales person per day, so a day can have several customers for them
          var cat = cats.treatment[i % cats.treatment.length];
          var draft = false;
          var status = k === perDay.length - 1 ? (j < 2 ? 'unpaid' : 'paid') : (i % 19 === 7 ? 'cancelled' : 'paid');
          var useCanula = i % 4 === 0, canulaAmount = useCanula ? 500 : 0;
          var canulaLines = useCanula ? [{ name: 'Canula 25G ยาว 50 mm', qty: 1, amount: 300 }, { name: 'Canula 22G', qty: 1, amount: 200 }] : [];
          var rid = 'R' + nextSeq_('RECORD', 6), label = 'บันทึกการรักษา ' + rid + ' · ' + patient.name;
          // 1 เวชภัณฑ์: supplies used.
          var supplyLines = [];
          if (i % 2 === 0) {
            var sp = supplies[i % supplies.length];
            supplyLines.push({ itemId: sp.id, name: sp.name, qty: 1 + (i % 3), unit: sp.unit });
          }
          // 2 หัตถการ: each line names its doctor; every sixth visit has a second one with another doctor.
          var procLines = [{ proc: procs[i % procs.length], doctor: doctor, amount: 1500 + (i % 8) * 900 }];
          if (i % 6 === 2) { var proc2 = procs[(i + 2) % procs.length]; procLines.push({ proc: proc2, doctor: doctors[(i + 1) % doctors.length], amount: proc2.price }); }
          var amount = round2_(procLines.reduce(function (sum, l) { return sum + l.amount; }, 0));
          // 3 ยา: จากใบเบิก (not charged; part may be left as ยาฝาก) / เบิกเพิ่ม (charged) / ใช้ยาฝาก / ถือยาเข้า / Messenger.
          var wdItems = [];
          if (fromWd) {
            var batch = 'WB' + nextSeq_('WITHDRAW_BATCH', 5);
            var wm = meds[i % meds.length], wq = 1 + (i % 4);
            var wdPlan = [[wm, wq, i % 5 === 0 ? 1 : wq]];
            if (i % 4 === 2) wdPlan.push([meds[(i + 1) % meds.length], 1, 1]);
            wdPlan.forEach(function (w) {
              var dep = round2_(w[1] - w[2]), expiry = '';
              if (dep > 0) {
                expiry = lotExpiry();
                lots.push({ id: 'DP' + nextSeq_('DEPOSIT', 5), date: date, patient: patient, itemId: w[0].id, name: w[0].name, expiry: expiry, qtyDeposited: dep, qty: dep, recordId: rid });
              }
              var wid = addWithdrawal(date, patient, salesP, w[0], w[1], rid, batch);
              wdItems.push({ withdrawId: wid, itemId: w[0].id, name: w[0].name, qtyDispensed: w[1], qtyUsed: w[2], qtyDeposit: dep, expiry: expiry });
            });
          }
          var medLines = [];
          if (i % 7 === 3) {
            var em = meds[(i + 2) % meds.length], eq = 1 + (i % 3), eup = seedTierPrice_(em, eq);
            // Every other one is รายการซื้อยาเพิ่มหน้าคลีนิค (คิดค่ายา): charged at the stock's ราคาขาย, flagged charged.
            var chargedMed = i % 14 === 3;
            medLines.push(chargedMed
              ? { itemId: em.id, name: em.name, qtyDispensed: eq, qtyUsed: eq, qtyDeposit: 0, expiry: '', charged: true, unitPrice: eup, amount: round2_(eq * eup), unitCost: em.cost }
              : { itemId: em.id, name: em.name, qtyDispensed: eq, qtyUsed: eq, qtyDeposit: 0, expiry: '', unitPrice: 0, amount: 0, unitCost: em.cost });
          }
          var depositUse = [];
          if (i % 2 === 1) {
            var lot = lots.filter(function (l) { return l.patient.id === patient.id && l.qty > 0 && l.expiry >= date && l.recordId !== rid; })
              .sort(function (x, y) { return x.expiry.localeCompare(y.expiry); })[0];
            if (lot) { lot.qty = round2_(lot.qty - 1); depositUse.push({ name: lot.name, qty: 1, itemId: lot.itemId, lots: [{ depositId: lot.id, qty: 1 }] }); }
          }
          // รายการจ่ายยาทาน: charged at the stock's ราคาขาย, counted in the ยอดรับชำระ.
          var oralLines = [];
          if (i % 6 === 1) { var om = orals[i % orals.length], oq = 1 + (i % 3); oralLines.push({ itemId: om.id, name: om.name, qty: oq, unit: om.unit, unitPrice: om.prices[0], amount: round2_(oq * om.prices[0]) }); }
          var oralAmount = round2_(oralLines.reduce(function (sum, x) { return sum + x.amount; }, 0));
          var ownMed = i % 9 === 4 ? [{ name: 'ยาที่ลูกค้านำมาเอง', qty: 1 }] : [];
          var messenger = i % 10 === 6 ? [{ name: wdItems.length ? wdItems[0].name : (medLines.length ? medLines[0].name : meds[0].name), qty: 1, note: 'Kerry TH' + (100000 + i) }] : [];
          // 4 รายการอื่นๆ: promo (its doctor is the next one, so those visits pay two doctors) / canula / other.
          var promo = activePromos.length && i % 5 === 0 ? activePromos[i % activePromos.length] : null;
          var promoDoctor = doctors[(i + 1) % doctors.length];
          var promoLine = promo ? { promoId: promo.id, name: promo.name, price: promo.price, doctorSharePct: promo.doctorSharePct, amount: promo.price, doctorId: promoDoctor.id } : null;
          var promoAmount = promo ? promo.price : 0;
          var otherAmount = i % 9 === 0 ? 300 : 0;
          var medicineAmount = round2_(medLines.reduce(function (sum, x) { return sum + x.amount; }, 0));
          var salesPct = salesP ? salesP.commissionRate : 0;
          var doctorEntries = procLines.map(function (l) { return { doctorId: l.doctor.id, base: l.amount }; });
          if (promo) doctorEntries.push({ doctorId: promoDoctor.id, extra: promo.price * promo.doctorSharePct / 100 });
          var row = {
            RID: rid, Date: date, PatientID: patient.id, PatientName: patient.name, Phone: patient.phone, ServiceCategory: cat, Detail: cat + ' - โปรแกรมมาตรฐาน',
            SalesID: salesP ? salesP.id : '', SalesName: salesP ? salesP.name : '',
            PaymentMethod: status === 'unpaid' ? '' : payMethods[i % payMethods.length], SuppliesJSON: JSON.stringify(supplyLines),
            UseCanula: useCanula ? 'TRUE' : 'FALSE', CanulaAmount: canulaAmount, CanulaJSON: JSON.stringify(canulaLines), Amount: amount, Status: status, ReceiptNo: '', Note: '',
            CreatedBy: actor, CreatedAt: nowStr_(), UpdatedAt: nowStr_(), OtherAmount: otherAmount,
            SalesCommissionPct: salesPct, SalesShare: round2_(amount * salesPct / 100),
            VisitStatus: draft ? 'dispensed' : 'completed',
            UsePromo: promo ? 'TRUE' : 'FALSE', PromoJSON: JSON.stringify(promoLine ? [promoLine] : []), PromoAmount: promoAmount,
            UseMedicine: medLines.length ? 'TRUE' : 'FALSE', MedicineJSON: JSON.stringify(medLines), MedicineAmount: medicineAmount,
            UseOralMed: oralLines.length ? 'TRUE' : 'FALSE', OralMedJSON: JSON.stringify(oralLines), OralMedAmount: oralAmount,
            UseOther: otherAmount ? 'TRUE' : 'FALSE', OtherItemsJSON: JSON.stringify(otherAmount ? [{ label: 'ค่าบริการอื่นๆ', amount: otherAmount }] : []),
            SupplierPayAmount: round2_(medicineAmount + otherAmount), OpdImageFileId: '',
            UseSupplies: supplyLines.length ? 'TRUE' : 'FALSE',
            UseOwnMed: ownMed.length ? 'TRUE' : 'FALSE', OwnMedJSON: JSON.stringify(ownMed),
            UseDepositUse: depositUse.length ? 'TRUE' : 'FALSE', DepositUseJSON: JSON.stringify(depositUse),
            UseMessenger: messenger.length ? 'TRUE' : 'FALSE', MessengerJSON: JSON.stringify(messenger),
            StockApplied: 'TRUE', WithdrawJSON: JSON.stringify(wdItems),
            UseProcedure: 'TRUE',
            ProcedureJSON: JSON.stringify(procLines.map(function (l) { return { procId: l.proc.id, name: l.proc.name, price: l.proc.price, amount: l.amount, doctorId: l.doctor.id, detail: l.proc.name + ' ตามแผนการรักษา' }; }))
          };
          applyDoctorShares_(row, buildDoctorShares_(doctorEntries, seedDoctorById));
          appendRow_(recSheet, CONFIG.HEADERS.Records, row);
          // Same stock effects saveRecord applies (supplies + ยาเบิกเพิ่ม; ยาจากใบเบิก left stock at its withdrawal), replayed in memory.
          var stockEff = recordStockEffects_(row);
          Object.keys(stockEff).forEach(function (id) {
            var item = stock.filter(function (s) { return s.id === id; })[0];
            if (item) move(date, item, 'record', 0, stockEff[id], label);
          });
          if (status !== 'cancelled') doctorEntries.forEach(function (e) { groups[e.doctorId + '|' + date] = { doctorId: e.doctorId, date: date }; });
        }
      });
      Object.keys(groups).forEach(function (key) { recomputeDoctorDayShares_(groups[key].doctorId, groups[key].date); });

      // The latest จัดยา-เบิกยา (yesterday = dispensed in advance, and today): each opened its record, with no
      // หัตถการ yet. Yesterday has two customers for the same sales person so รายงานรายเซลล์ shows a stacked day.
      // The status pass below keeps the first two at จัดยาแล้ว, and moves the rest to รอพบแพทย์ / พบแพทย์.
      // The first one was typed in at the counter (name not in the list): its patient is "ชื่อ (รออัพเดท)" until the history is filled in.
      var pendingId = 'P' + nextSeq_('PATIENT', 5);
      var pendingPatient = { id: pendingId, hn: '', name: 'คุณสมศรี ' + PENDING_NAME_TAG_, phone: '' };
      appendRow_(getSheet_(CONFIG.SHEETS.PATIENTS), CONFIG.HEADERS.Patients, {
        ID: pendingId, HN: '', FirstName: pendingPatient.name, LastName: '', Phone: '', DOB: '', Gender: '', Email: '', PhotoFileId: '',
        JoinedDate: dayStr(1), CreatedAt: nowStr_(), UpdatedAt: nowStr_(), Nickname: ''
      });
      [[1, 0, 0, [[0, 2], [2, 3]]], [1, 0, 5, [[2, 5], [3, 1]]], [1, 1, 3, [[1, 1]]], [0, 0, 7, [[4, 2]]], [0, 1, 2, [[0, 1]]], [0, 1, 6, [[3, 2]]]].forEach(function (w, wi) {
        var date = dayStr(w[0]), salesP = sales[w[1]], patient = wi === 0 ? pendingPatient : patients[w[2]];
        var rid = 'R' + nextSeq_('RECORD', 6), batch = 'WB' + nextSeq_('WITHDRAW_BATCH', 5);
        var items = w[3].map(function (pair) {
          var m = meds[pair[0]], q = pair[1];
          return { withdrawId: addWithdrawal(date, patient, salesP, m, q, rid, batch), itemId: m.id, name: m.name, qtyDispensed: q, qtyUsed: q, qtyDeposit: 0, expiry: '' };
        });
        var rec = {
          RID: rid, Date: date, PatientID: patient.id, PatientName: patient.name, Phone: patient.phone,
          SalesID: salesP.id, SalesName: salesP.name, Status: 'unpaid', PaymentMethod: '', ReceiptNo: '', Note: '',
          CreatedBy: actor, CreatedAt: nowStr_(), UpdatedAt: nowStr_(),
          SalesCommissionPct: salesP.commissionRate, SalesShare: 0, VisitStatus: 'dispensed', StockApplied: 'TRUE', WithdrawJSON: JSON.stringify(items)
        };
        applyDoctorShares_(rec, []);
        appendRow_(recSheet, CONFIG.HEADERS.Records, rec);
        i++;
      });

      // ตัวอย่างครบวงจร: ลูกค้า 1 ราย 1 วัน (วันนี้) — ประวัติคนไข้ (HN) > จัดยา-เบิกยา (โบท็อกซ์ 3 ใช้ 1 ฝาก 2 + วิตามินซี 2)
      // > บันทึกการรักษา (เวชภัณฑ์ · หัตถการ 2 รายการต่างแพทย์ · ยาจากใบเบิก + ยาเบิกเพิ่ม · โปร / Canula / อื่นๆ)
      // > เสร็จสิ้น > ชำระเงินสด + ออกใบเสร็จ. Shows in every list, the day summary and both reports.
      (function () {
        var date = dayStr(0), salesP = sales[0], dA = doctors[0], dB = doctors[1];
        var pid = 'P' + nextSeq_('PATIENT', 5), hn = 'HN' + String(1001 + patients.length);
        var patient = { id: pid, hn: hn, name: 'ตัวอย่าง ครบวงจร', phone: '0800000011' };
        appendRow_(getSheet_(CONFIG.SHEETS.PATIENTS), CONFIG.HEADERS.Patients, {
          ID: pid, HN: hn, FirstName: 'ตัวอย่าง', LastName: 'ครบวงจร', Phone: patient.phone, DOB: '', Gender: 'female', Email: '', PhotoFileId: '',
          JoinedDate: date, CreatedAt: nowStr_(), UpdatedAt: nowStr_(), Nickname: 'เดโม'
        });
        patients.push(patient);
        var rid = 'R' + nextSeq_('RECORD', 6), batch = 'WB' + nextSeq_('WITHDRAW_BATCH', 5), label = 'บันทึกการรักษา ' + rid + ' · ' + patient.name;
        // จัดยา-เบิกยา → ยาจากใบเบิก (ฝาก 2 ขวดไว้ใช้ครั้งหน้า)
        var botox = meds[0], vitC = meds[2], cream = meds[3], lotExp = expiryAfter(180);
        lots.push({ id: 'DP' + nextSeq_('DEPOSIT', 5), date: date, patient: patient, itemId: botox.id, name: botox.name, expiry: lotExp, qtyDeposited: 2, qty: 2, recordId: rid });
        var wdItems = [
          { withdrawId: addWithdrawal(date, patient, salesP, botox, 3, rid, batch), itemId: botox.id, name: botox.name, qtyDispensed: 3, qtyUsed: 1, qtyDeposit: 2, expiry: lotExp },
          { withdrawId: addWithdrawal(date, patient, salesP, vitC, 2, rid, batch), itemId: vitC.id, name: vitC.name, qtyDispensed: 2, qtyUsed: 2, qtyDeposit: 0, expiry: '' }
        ];
        // บันทึกการรักษา
        var procLines = [{ proc: procs[0], doctor: dA, amount: procs[0].price }, { proc: procs[1], doctor: dB, amount: procs[1].price }];
        var amount = round2_(procLines.reduce(function (sum, l) { return sum + l.amount; }, 0));
        var promo = activePromos[2], promoLine = { promoId: promo.id, name: promo.name, price: promo.price, doctorSharePct: promo.doctorSharePct, amount: promo.price, doctorId: dB.id };
        var oralDemo = orals[0], oralDemoLines = [{ itemId: oralDemo.id, name: oralDemo.name, qty: 2, unit: oralDemo.unit, unitPrice: oralDemo.prices[0], amount: round2_(2 * oralDemo.prices[0]) }];
        var creamPrice = seedTierPrice_(cream, 1); // ซื้อยาเพิ่มหน้าคลีนิค (คิดค่ายา)
        var extraLines = [{ itemId: cream.id, name: cream.name, qtyDispensed: 1, qtyUsed: 1, qtyDeposit: 0, expiry: '', charged: true, unitPrice: creamPrice, amount: creamPrice, unitCost: cream.cost }];
        var supplyLines = [{ itemId: supplies[0].id, name: supplies[0].name, qty: 2, unit: supplies[0].unit }];
        var salesPct = salesP.commissionRate;
        var row = {
          RID: rid, Date: date, PatientID: pid, PatientName: patient.name, Phone: patient.phone, ServiceCategory: '', Detail: 'ตัวอย่างครบวงจร: โบท็อกซ์ + ฟิลเลอร์ (แพทย์ 2 ท่าน) พร้อมโปรโมชัน',
          SalesID: salesP.id, SalesName: salesP.name, PaymentMethod: 'เงินสด', SuppliesJSON: JSON.stringify(supplyLines),
          UseCanula: 'TRUE', CanulaAmount: 500, CanulaJSON: JSON.stringify([{ name: 'Canula 25G ยาว 50 mm', qty: 2, amount: 500 }]), Amount: amount, Status: 'paid', ReceiptNo: generateReceiptNo_(), Note: '',
          CreatedBy: actor, CreatedAt: nowStr_(), UpdatedAt: nowStr_(), OtherAmount: 300,
          SalesCommissionPct: salesPct, SalesShare: round2_(amount * salesPct / 100), VisitStatus: 'completed',
          UsePromo: 'TRUE', PromoJSON: JSON.stringify([promoLine]), PromoAmount: promo.price,
          UseMedicine: 'TRUE', MedicineJSON: JSON.stringify(extraLines), MedicineAmount: extraLines[0].amount,
          UseOralMed: 'TRUE', OralMedJSON: JSON.stringify(oralDemoLines), OralMedAmount: oralDemoLines[0].amount,
          UseOther: 'TRUE', OtherItemsJSON: JSON.stringify([{ label: 'ค่าบริการอื่นๆ', amount: 300 }]),
          SupplierPayAmount: round2_(extraLines[0].amount + 300), OpdImageFileId: '',
          UseSupplies: 'TRUE', UseOwnMed: 'FALSE', OwnMedJSON: '[]', UseDepositUse: 'FALSE', DepositUseJSON: '[]', UseMessenger: 'FALSE', MessengerJSON: '[]',
          StockApplied: 'TRUE', WithdrawJSON: JSON.stringify(wdItems), UseProcedure: 'TRUE',
          ProcedureJSON: JSON.stringify(procLines.map(function (l, li) { return { procId: l.proc.id, name: l.proc.name, price: l.proc.price, amount: l.amount, doctorId: l.doctor.id, detail: li === 0 ? 'ฉีดบริเวณหน้าผากและหางตา' : 'บริเวณร่องแก้ม' }; }))
        };
        var entries = procLines.map(function (l) { return { doctorId: l.doctor.id, base: l.amount }; }).concat([{ doctorId: dB.id, extra: promo.price * promo.doctorSharePct / 100 }]);
        applyDoctorShares_(row, buildDoctorShares_(entries, seedDoctorById));
        appendRow_(recSheet, CONFIG.HEADERS.Records, row);
        var stockEff = recordStockEffects_(row);
        Object.keys(stockEff).forEach(function (id) {
          var item = stock.filter(function (st) { return st.id === id; })[0];
          if (item) move(date, item, 'record', 0, stockEff[id], label);
        });
        [dA.id, dB.id].forEach(function (id) { recomputeDoctorDayShares_(id, date); });
        i++;
      })();

      // ตัวอย่างยาครบทุกแบบ: ลูกค้า 1 ราย 2 ครั้ง. ครั้งแรก (30 วันก่อน) ฝากโบท็อกซ์ 2 ขวดไว้; วันนี้ (ชำระเงินสด + ใบเสร็จ) บันทึกการรักษา
      // มียาครบทุกกลุ่ม — คิดเงิน: ซื้อยาเพิ่มหน้าคลีนิค (คิดค่ายา) + รายการจ่ายยาทาน; ไม่คิดเงิน: ยาจากใบเบิก, ซื้อยาเพิ่มหน้าคลีนิค
      // (ไม่คิดค่ายา), ใช้ยาฝาก (จากครั้งแรก), ถือยาเข้า, Messenger/พัสดุ.
      (function () {
        var salesP = sales[1 % sales.length], dA = doctors[0];
        var pid = 'P' + nextSeq_('PATIENT', 5), hn = 'HN' + String(1001 + patients.length);
        var patient = { id: pid, hn: hn, name: 'ตัวอย่าง ยาครบทุกแบบ', phone: '0800000012' };
        appendRow_(getSheet_(CONFIG.SHEETS.PATIENTS), CONFIG.HEADERS.Patients, {
          ID: pid, HN: hn, FirstName: 'ตัวอย่าง', LastName: 'ยาครบทุกแบบ', Phone: patient.phone, DOB: '', Gender: 'male', Email: '', PhotoFileId: '',
          JoinedDate: dayStr(30), CreatedAt: nowStr_(), UpdatedAt: nowStr_(), Nickname: 'เดโมยา'
        });
        patients.push(patient);
        var botox = meds[0], vitC = meds[2], cream = meds[3], para = meds[4];
        var write = function (date, rid, row, entries) {
          applyDoctorShares_(row, buildDoctorShares_(entries, seedDoctorById));
          appendRow_(recSheet, CONFIG.HEADERS.Records, row);
          var eff = recordStockEffects_(row);
          Object.keys(eff).forEach(function (id) {
            var item = stock.filter(function (st) { return st.id === id; })[0];
            if (item) move(date, item, 'record', 0, eff[id], 'บันทึกการรักษา ' + rid + ' · ' + patient.name);
          });
          entries.forEach(function (e) { recomputeDoctorDayShares_(e.doctorId, date); });
          i++;
        };
        var base = function (rid, date, extra) {
          return Object.assign({
            RID: rid, Date: date, PatientID: pid, PatientName: patient.name, Phone: patient.phone, ServiceCategory: '',
            SalesID: salesP.id, SalesName: salesP.name, PaymentMethod: 'เงินสด', SuppliesJSON: '[]', UseCanula: 'FALSE', CanulaAmount: 0, CanulaJSON: '[]',
            Status: 'paid', ReceiptNo: '', Note: '', CreatedBy: actor, CreatedAt: nowStr_(), UpdatedAt: nowStr_(), OtherAmount: 0,
            SalesCommissionPct: salesP.commissionRate, VisitStatus: 'completed', UsePromo: 'FALSE', PromoJSON: '[]', PromoAmount: 0,
            UseOther: 'FALSE', OtherItemsJSON: '[]', OpdImageFileId: '', UseSupplies: 'FALSE', StockApplied: 'TRUE', UseProcedure: 'TRUE',
            UseMedicine: 'FALSE', MedicineJSON: '[]', MedicineAmount: 0, UseOralMed: 'FALSE', OralMedJSON: '[]', OralMedAmount: 0,
            UseOwnMed: 'FALSE', OwnMedJSON: '[]', UseDepositUse: 'FALSE', DepositUseJSON: '[]', UseMessenger: 'FALSE', MessengerJSON: '[]', WithdrawJSON: '[]'
          }, extra);
        };
        var procLine = function (p, detail) { return { procId: p.id, name: p.name, price: p.price, amount: p.price, doctorId: dA.id, detail: detail }; };

        // ครั้งแรก: จัดยา-เบิกยา โบท็อกซ์ 3 ใช้ 1 ฝาก 2
        var d1 = dayStr(30), rid1 = 'R' + nextSeq_('RECORD', 6), batch1 = 'WB' + nextSeq_('WITHDRAW_BATCH', 5), lotExp = expiryAfter(150);
        var lot = { id: 'DP' + nextSeq_('DEPOSIT', 5), date: d1, patient: patient, itemId: botox.id, name: botox.name, expiry: lotExp, qtyDeposited: 2, qty: 2, recordId: rid1 };
        lots.push(lot);
        var wd1 = [{ withdrawId: addWithdrawal(d1, patient, salesP, botox, 3, rid1, batch1), itemId: botox.id, name: botox.name, qtyDispensed: 3, qtyUsed: 1, qtyDeposit: 2, expiry: lotExp }];
        var p1 = procLine(procs[0], 'ฉีดกรามทั้งสองข้าง');
        write(d1, rid1, base(rid1, d1, {
          Detail: 'ครั้งแรก: ฝากโบท็อกซ์ 2 ขวดไว้ใช้ครั้งถัดไป', Amount: p1.amount, SalesShare: round2_(p1.amount * salesP.commissionRate / 100),
          WithdrawJSON: JSON.stringify(wd1), ProcedureJSON: JSON.stringify([p1])
        }), [{ doctorId: dA.id, base: p1.amount }]);

        // วันนี้: ยาครบทุกแบบ
        var d2 = dayStr(0), rid2 = 'R' + nextSeq_('RECORD', 6), batch2 = 'WB' + nextSeq_('WITHDRAW_BATCH', 5);
        // ไม่คิดเงิน — ยาจากใบเบิก (วิตามินซีฉีด 2 ใช้ 2)
        var wd2 = [{ withdrawId: addWithdrawal(d2, patient, salesP, vitC, 2, rid2, batch2), itemId: vitC.id, name: vitC.name, qtyDispensed: 2, qtyUsed: 2, qtyDeposit: 0, expiry: '' }];
        // ซื้อยาเพิ่มหน้าคลีนิค: (ไม่คิดค่ายา) ยาแก้ปวด 1 แผง + (คิดค่ายา) ครีมทาแผล 2 หลอด ที่ราคาขายตามจำนวน
        var creamPrice = seedTierPrice_(cream, 2);
        var medLines = [
          { itemId: para.id, name: para.name, qtyDispensed: 1, qtyUsed: 1, qtyDeposit: 0, expiry: '', unitPrice: 0, amount: 0, unitCost: para.cost },
          { itemId: cream.id, name: cream.name, qtyDispensed: 2, qtyUsed: 2, qtyDeposit: 0, expiry: '', charged: true, unitPrice: creamPrice, amount: round2_(2 * creamPrice), unitCost: cream.cost }
        ];
        // คิดเงิน — รายการจ่ายยาทาน 2 รายการ
        var oralLines = [orals[0], orals[2]].map(function (o, k) { var q = k === 0 ? 2 : 1; return { itemId: o.id, name: o.name, qty: q, unit: o.unit, unitPrice: o.prices[0], amount: round2_(q * o.prices[0]) }; });
        // ไม่คิดเงิน — ใช้ยาฝากจากครั้งแรก 1 ขวด, ถือยาเข้า, ส่งยาทาง Messenger
        lot.qty = round2_(lot.qty - 1);
        var depUse = [{ name: botox.name, qty: 1, itemId: botox.id, lots: [{ depositId: lot.id, qty: 1 }] }];
        var ownMed = [{ name: 'ครีมกันแดดที่ลูกค้านำมาเอง', qty: 1 }];
        var messenger = [{ name: orals[1].name, qty: 1, note: 'ส่ง Kerry TH900012' }];
        var p2 = procLine(procs[0], 'ฉีดโบท็อกซ์จากยาฝาก บริเวณกราม'), p3 = procLine(procs[2 % procs.length], 'ทำเลเซอร์ทั่วหน้า');
        var amount2 = round2_(p2.amount + p3.amount);
        var medAmount = round2_(medLines.reduce(function (sum, l) { return sum + l.amount; }, 0));
        var oralAmount = round2_(oralLines.reduce(function (sum, l) { return sum + l.amount; }, 0));
        write(d2, rid2, base(rid2, d2, {
          Detail: 'ตัวอย่างยาครบทุกแบบ: ยาคิดเงิน (ซื้อยาเพิ่ม คิดค่ายา, ยาทาน) และยาไม่คิดเงิน (ใบเบิก, ซื้อยาเพิ่ม ไม่คิดค่ายา, ยาฝาก, ถือยาเข้า, Messenger)',
          Amount: amount2, SalesShare: round2_(amount2 * salesP.commissionRate / 100), ReceiptNo: generateReceiptNo_(),
          ProcedureJSON: JSON.stringify([p2, p3]), WithdrawJSON: JSON.stringify(wd2),
          UseMedicine: 'TRUE', MedicineJSON: JSON.stringify(medLines), MedicineAmount: medAmount, SupplierPayAmount: medAmount,
          UseOralMed: 'TRUE', OralMedJSON: JSON.stringify(oralLines), OralMedAmount: oralAmount,
          UseDepositUse: 'TRUE', DepositUseJSON: JSON.stringify(depUse),
          UseOwnMed: 'TRUE', OwnMedJSON: JSON.stringify(ownMed),
          UseMessenger: 'TRUE', MessengerJSON: JSON.stringify(messenger)
        }), [{ doctorId: dA.id, base: amount2 }]);
      })();

      // สถานะการรักษา — every customer who has arrived has the day's queue number (in order of arrival: those already
      // treated first, then the ones with a doctor, then the ones still waiting) and the doctor they saw / are seeing.
      (function () {
        var today = dayStr(0);
        var docOf = function (r, fallback) { var sh = recordDoctorShares_(r)[0]; return sh ? { id: sh.doctorId, name: sh.name } : fallback; };
        var rows = getRows_(recSheet).filter(function (r) { return r.Status !== 'cancelled'; });
        var write = function (r) { updateRowByIndex_(recSheet, CONFIG.HEADERS.Records, r._row, r); };
        // 1. visits that are already treated (รอชำระเงิน / เสร็จสิ้น), oldest day first
        rows.filter(function (r) { return visitStatusOf_(r) !== 'dispensed'; })
          .sort(function (a, b) { return dateStr_(a.Date).localeCompare(dateStr_(b.Date)) || String(a.RID).localeCompare(String(b.RID)); })
          .forEach(function (r) {
            var day = dateStr_(r.Date), d = docOf(r, { id: doctors[0].id, name: doctors[0].name });
            r.QueueDate = day; r.QueueDoctor = nextQueueNo_('D', day); r.VisitDoctorID = d.id; r.VisitDoctorName = d.name;
            write(r);
          });
        // 2. the จัดยา-เบิกยา-only records, in order: two stay จัดยาแล้ว (dispensed in advance), two have arrived (รอพบแพทย์),
        //    two are with a doctor (different doctors). The ones with a doctor arrived first, so they take the earlier numbers.
        var onlyDispensed = rows.filter(function (r) { return visitStatusOf_(r) === 'dispensed' && !recordHasContent_(r); })
          .sort(function (a, b) { return String(a.RID).localeCompare(String(b.RID)); });
        var plan = [['dispensed'], ['dispensed'], ['waiting_doctor'], ['waiting_doctor'], ['with_doctor', 0], ['with_doctor', 1]];
        var order = [4, 5, 2, 3];
        order.forEach(function (idx) {
          var r = onlyDispensed[idx], step = plan[idx];
          if (!r) return;
          var doc = step[1] != null ? { id: doctors[step[1] % doctors.length].id, name: doctors[step[1] % doctors.length].name } : null;
          r.VisitStatus = 'dispensed';
          applyVisitStatus_(r, step[0], today, true, doc);
          write(r);
        });
      })();

      events.sort(function (a, b) { return a.date < b.date ? -1 : a.date > b.date ? 1 : a.idx - b.idx; });
      events.forEach(function (ev) { ev.run(); });
      // One ปรับยอดคงเหลือ today after a physical count (2 short), so the history shows that kind too.
      var counted = supplies[1];
      apply(dayStr(0), counted, 'set', 0, Math.min(2, Math.max(0, counted.qty)), 'นับสต๊อกจริง');

      // ---- Write Stock (final Qty), StockLog, MedDeposits ----
      var stockSheet = getSheet_(CONFIG.SHEETS.STOCK);
      stock.forEach(function (s) {
        var row = { ItemID: s.id, Name: s.name, Category: s.category, Unit: s.unit, Qty: s.qty, CreatedAt: nowStr_(), UpdatedAt: nowStr_(), StockType: s.type, MinQty: s.min, CostPrice: s.cost != null ? s.cost : '' };
        STOCK_PRICE_TIERS.forEach(function (t, k) { row['Price' + t] = s.prices ? s.prices[k] : ''; });
        appendRow_(stockSheet, CONFIG.HEADERS.Stock, row);
      });
      var logIds = nextSeqBlock_('STOCKLOG', logs.length, 5);
      var logSheet = getSheet_(CONFIG.SHEETS.STOCK_LOG);
      var logValues = logs.map(function (lg, k) {
        lg.LogID = 'SL' + logIds[k];
        return CONFIG.HEADERS.StockLog.map(function (h) { return lg[h] !== undefined ? lg[h] : ''; });
      });
      if (logValues.length) {
        var firstLogRow = logSheet.getLastRow() + 1;
        resetNumberFormats_(logSheet, CONFIG.HEADERS.StockLog, firstLogRow, logValues.length);
        logSheet.getRange(firstLogRow, 1, logValues.length, CONFIG.HEADERS.StockLog.length).setValues(logValues);
      }
      var depSheet = getSheet_(CONFIG.SHEETS.MED_DEPOSITS);
      lots.forEach(function (l) {
        appendRow_(depSheet, CONFIG.HEADERS.MedDeposits, {
          DepositID: l.id, Date: l.date, PatientID: l.patient.id, PatientName: l.patient.name, ItemName: l.name, QtyRemaining: l.qty,
          CreatedBy: actor, ItemID: l.itemId, UpdatedAt: nowStr_(), ExpiryDate: l.expiry, RecordID: l.recordId, QtyDeposited: l.qtyDeposited
        });
      });

      // ---- Expenses ----
      var expSheet = getSheet_(CONFIG.SHEETS.EXPENSES);
      for (var e = 0; e < 15; e++) {
        var ec = cats.expense[e % cats.expense.length];
        appendRow_(expSheet, CONFIG.HEADERS.Expenses, { ExpID: 'EX' + nextSeq_('EXPENSE', 5), Date: dayStr(e * 6), Name: ec + ' ประจำเดือน', Category: ec, Amount: 2000 + (e % 5) * 1500, CreatedBy: actor, CreatedAt: nowStr_() });
      }

      // ---- ClinicSettings: only fill keys that are still empty (real company info and logo are kept) ----
      var settingsSheet = getSheet_(CONFIG.SHEETS.CLINIC_SETTINGS);
      var have = {};
      getRows_(settingsSheet).forEach(function (r) { if (String(r.Value || '') !== '') have[r.Key] = true; });
      var defaults = { name: 'VALIN CLINIC', address: '123 ถนนสุขุมวิท แขวงคลองตัน เขตคลองเตย กรุงเทพฯ 10110', phone: '02-123-4567', taxId: '0-1055-12345-67-8', email: 'contact@valinclinic.com' };
      Object.keys(defaults).forEach(function (key) {
        if (!have[key]) appendRow_(settingsSheet, CONFIG.HEADERS.ClinicSettings, { Key: key, Value: defaults[key] });
      });

      invalidateMasterCache_();
      invalidatePatientsLiteCache_();
      invalidateStockLiteCache_();
      bumpDashboardCacheVersion_();
      return { ok: true, message: 'สร้างข้อมูลตัวอย่างเรียบร้อยแล้ว', data: { records: i, stockItems: stock.length, stockLogs: logs.length, deposits: lots.length } };
    } finally {
      lock.releaseLock();
    }
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
  // ยอดสุทธิ = จำนวนเงินทั้งหมด - ส่วนแบ่งแพทย์ - ส่วนแบ่งเซลล์ (== s.clinicNet;
  // there is no supplier share any more).
  lines.push('ยอดสุทธิ: ' + moneyFmt_(s.clinicNet) + ' บาท');
  lines.push('จำนวนเงินทั้งหมด: ' + moneyFmt_(s.totalCollected) + ' บาท');
  lines.push('- ส่วนแบ่งแพทย์: ' + moneyFmt_(s.doctorTotal) + ' บาท');
  lines.push('- ส่วนแบ่งเซลล์: ' + moneyFmt_(s.salesTotal) + ' บาท');
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

/* ======================================================================
   BACKUP — ย้ายข้อมูลที่จบแล้วไปไฟล์ใหม่ (admin only)
   The finished part of the big sheets is copied into a NEW spreadsheet saved in the Drive folder "back up data" (inside the
   root folder), checked, and then removed from the main sheets so they stay fast. What the dashboard needs from those days
   is first written to DailySummary (see CONFIG.HEADERS.DailySummary), and every backup is listed in Backups.
   "จบแล้ว" before the cutoff date: paid / cancelled treatment records and the จัดยา-เบิกยา lines of them, ค่าใช้จ่าย,
   ใบสั่งซื้อ (+ lines), the AuditLog and the StockLog (each item's latest entry stays: it is the item's balance).
   Not moved: unfinished records, patients, stock, ยาฝาก lots, settings.
   ====================================================================== */
var BACKUP_FOLDER_NAME_ = 'back up data';
// [sheet name, headers key] in the order they are written to the backup file
var BACKUP_SHEETS_ = [
  ['Records', 'RECORDS'], ['MedWithdrawals', 'MED_WITHDRAWALS'], ['StockLog', 'STOCK_LOG'], ['Expenses', 'EXPENSES'],
  ['Orders', 'ORDERS'], ['OrderLines', 'ORDER_LINES'], ['AuditLog', 'AUDIT_LOG']
];

/** What would move for `cutoff` (rows dated before it): {move:{sheet:[rows]}, keep:{sheet:[rows]}, from, to}. Reads only. */
function backupPlan_(cutoff) {
  var S = CONFIG.SHEETS;
  var records = getRows_(getSheet_(S.RECORDS)), wds = getRows_(getSheet_(S.MED_WITHDRAWALS));
  var logs = getRows_(getSheet_(S.STOCK_LOG)), expenses = getRows_(getSheet_(S.EXPENSES));
  var orders = getRows_(getSheet_(S.ORDERS)), lines = getRows_(getSheet_(S.ORDER_LINES)), audit = getRows_(getSheet_(S.AUDIT_LOG));
  var plan = { move: {}, keep: {}, cutoff: cutoff };
  var split = function (name, rows, fn) {
    plan.move[name] = []; plan.keep[name] = [];
    rows.forEach(function (r) { (fn(r) ? plan.move[name] : plan.keep[name]).push(r); });
  };
  var rids = {}, movedRid = {};
  records.forEach(function (r) { rids[r.RID] = true; });
  split('Records', records, function (r) {
    var done = (r.Status === 'paid' || r.Status === 'cancelled') && dateStr_(r.Date) < cutoff;
    if (done) movedRid[r.RID] = true;
    return done;
  });
  // a line goes with its record; a line whose record is gone (or never had one) goes when it is old enough
  split('MedWithdrawals', wds, function (w) {
    if (w.RecordID && rids[w.RecordID]) return !!movedRid[w.RecordID];
    return dateStr_(w.Date) < cutoff;
  });
  var latest = {};
  logs.forEach(function (l) {
    var cur = latest[l.ItemID];
    if (!cur || dateStr_(l.Date) > dateStr_(cur.Date) || (dateStr_(l.Date) === dateStr_(cur.Date) && String(l.CreatedAt) > String(cur.CreatedAt))) latest[l.ItemID] = l;
  });
  split('StockLog', logs, function (l) { return dateStr_(l.Date) < cutoff && latest[l.ItemID] !== l; });
  split('Expenses', expenses, function (e) { return dateStr_(e.Date) < cutoff; });
  var movedPo = {};
  split('Orders', orders, function (o) { var m = dateStr_(o.Date) < cutoff; if (m) movedPo[o.PoID] = true; return m; });
  split('OrderLines', lines, function (l) { return !!movedPo[l.PoID]; });
  split('AuditLog', audit, function (a) { return dateStr_(a.Timestamp) < cutoff; });
  var from = '', to = '';
  [['Records', 'Date'], ['MedWithdrawals', 'Date'], ['StockLog', 'Date'], ['Expenses', 'Date'], ['Orders', 'Date'], ['AuditLog', 'Timestamp']].forEach(function (p) {
    plan.move[p[0]].forEach(function (r) {
      var d = dateStr_(r[p[1]]);
      if (!d) return;
      if (!from || d < from) from = d;
      if (!to || d > to) to = d;
    });
  });
  plan.from = from; plan.to = to;
  return plan;
}

function backupCounts_(plan) {
  var out = {};
  BACKUP_SHEETS_.forEach(function (b) { out[b[0]] = { move: plan.move[b[0]].length, keep: plan.keep[b[0]].length }; });
  // records that stay because they are not finished (not paid / cancelled) or not old enough
  out.RecordsUnfinished = plan.keep.Records.filter(function (r) { return r.Status !== 'paid' && r.Status !== 'cancelled'; }).length;
  return out;
}

/** Admin: what a backup up to `cutoff` (exclusive) would move, per kind of data. */
function getBackupPreview(token, cutoff) {
  return safeCall_(function () {
    requireAuth_(token, 'admin');
    requireDateStr_(cutoff, 'วันที่ตัดยอด');
    if (cutoff > todayStr_()) throw ValidationError_('วันที่ตัดยอดต้องไม่เกินวันนี้');
    var plan = backupPlan_(cutoff);
    return { ok: true, data: { cutoff: cutoff, from: plan.from, to: plan.to, counts: backupCounts_(plan) } };
  });
}

function mapBackupOut_(b) {
  var counts = {};
  try { counts = JSON.parse(b.CountsJSON || '{}'); } catch (e) { counts = {}; }
  return {
    id: b.BackupID, createdAt: String(b.CreatedAt), createdBy: b.CreatedBy, from: dateStr_(b.FromDate), to: dateStr_(b.ToDate), cutoff: dateStr_(b.CutoffDate),
    fileName: b.FileName, spreadsheetId: b.SpreadsheetID, url: b.SpreadsheetUrl, status: b.Status, counts: counts
  };
}

/** Admin: every backup file made so far, newest first (name, sheet id, link, counts). */
function listBackups(token) {
  return safeCall_(function () {
    requireAuth_(token, 'admin');
    var rows = getRows_(getSheet_(CONFIG.SHEETS.BACKUPS)).map(mapBackupOut_);
    rows.sort(function (a, b) { return b.createdAt.localeCompare(a.createdAt); });
    return { ok: true, data: rows };
  });
}

/** The days moved to backups, summed for the dashboard: {m:{metric:total}, doctor, sales, promo, expense:{category:amount}, expenseTotal, patients:{id:true}}. */
function archiveSummary_(dateFrom, dateTo) {
  var out = { m: { customerRevenue: 0, proc: 0, promo: 0, canula: 0, drug: 0, salesIncome: 0, doctorShareTotal: 0, salesShareTotal: 0, orders: 0 },
    doctor: {}, sales: {}, promo: {}, expense: {}, expenseTotal: 0, patients: {} };
  getRows_(getSheet_(CONFIG.SHEETS.DAILY_SUMMARY)).forEach(function (r) {
    var d = dateStr_(r.Date);
    if (d < dateFrom || d > dateTo) return;
    var amt = Number(r.Amount) || 0, key = String(r.Key || '');
    switch (r.Metric) {
      case 'doctorShare': out.doctor[key] = (out.doctor[key] || 0) + amt; break;
      case 'salesShare': out.sales[key] = (out.sales[key] || 0) + amt; break;
      case 'promoLine': out.promo[key] = (out.promo[key] || 0) + amt; break;
      case 'expense': out.expense[key] = (out.expense[key] || 0) + amt; out.expenseTotal += amt; break;
      case 'patient': out.patients[key] = true; break;
      default: if (out.m[r.Metric] !== undefined) out.m[r.Metric] += amt;
    }
  });
  return out;
}

/** Archived revenue / expense per month {yyyy-MM: {revenue, expense}} for the dashboard's monthly trend. */
function archiveMonthly_() {
  var out = {};
  getRows_(getSheet_(CONFIG.SHEETS.DAILY_SUMMARY)).forEach(function (r) {
    var m = monthKey_(dateStr_(r.Date)), amt = Number(r.Amount) || 0;
    var o = out[m] || (out[m] = { revenue: 0, expense: 0 });
    if (r.Metric === 'customerRevenue' || r.Metric === 'salesIncome') o.revenue += amt;
    else if (r.Metric === 'expense' || r.Metric === 'orders' || r.Metric === 'doctorShareTotal' || r.Metric === 'salesShareTotal') o.expense += amt;
  });
  return out;
}

/** DailySummary rows (objects) for what the plan moves — only paid records count as revenue, like the dashboard. */
function backupSummaryRows_(plan, backupId) {
  var acc = {}, order = [];
  var add = function (date, metric, key, amount, count) {
    var k = date + '|' + metric + '|' + key;
    if (!acc[k]) { acc[k] = { Date: date, Metric: metric, Key: key, Amount: 0, Count: 0, BackupID: backupId }; order.push(k); }
    acc[k].Amount = round2_(acc[k].Amount + amount);
    acc[k].Count += count || 0;
  };
  var paid = plan.move.Records.filter(function (r) { return r.Status === 'paid'; });
  paid.forEach(function (r) {
    var d = dateStr_(r.Date);
    add(d, 'customerRevenue', '', recordCollected_(r));
    add(d, 'proc', '', Number(r.Amount) || 0);
    add(d, 'promo', '', Number(r.PromoAmount) || 0);
    add(d, 'canula', '', isTrue_(r.UseCanula) ? (Number(r.CanulaAmount) || 0) : 0);
    add(d, 'drug', '', (Number(r.MedicineAmount) || 0) + (Number(r.OralMedAmount) || 0));
    add(d, 'doctorShareTotal', '', Number(r.DoctorShare) || 0);
    add(d, 'salesShareTotal', '', Number(r.SalesShare) || 0);
    if (r.SalesID) add(d, 'salesShare', r.SalesName || '', Number(r.SalesShare) || 0);
    recordDoctorShares_(r).forEach(function (x) { add(d, 'doctorShare', x.name || '', Number(x.share) || 0); });
    parseLineItems_(r.PromoJSON).forEach(function (p) { add(d, 'promoLine', p.name || 'ไม่ระบุ', Number(p.amount) || 0); });
    if (r.PatientID && !acc[d + '|patient|' + r.PatientID]) add(d, 'patient', r.PatientID, 0, 1);
  });
  plan.move.Expenses.forEach(function (e) { add(dateStr_(e.Date), 'expense', e.Category || '', Number(e.Amount) || 0); });
  plan.move.Orders.forEach(function (o) { add(dateStr_(o.Date), 'orders', '', Number(o.TotalAmount) || 0); });
  salesReportRows_(function () { return true; }, { withdrawals: plan.move.MedWithdrawals, records: plan.move.Records }).forEach(function (o) {
    add(o.date, 'salesIncome', '', o.costAmount);
  });
  return order.map(function (k) { return acc[k]; });
}

/** Rows (objects from getRows_) -> 2-D values in header order; numbers read back as dates are written back as numbers. */
function backupValues_(sheetName, headers, rows) {
  var numCols = CONFIG.NUMBER_COLUMNS[sheetName] || [];
  return rows.map(function (o) {
    return headers.map(function (h) {
      var v = o[h] !== undefined && o[h] !== null ? o[h] : '';
      return isDate_(v) && numCols.indexOf(h) > -1 ? numOrNull_(v) : v;
    });
  });
}

/** Writes values below the header row of `sh` in blocks (a big block in one call can time out). */
function backupWriteBlocks_(sh, headers, values) {
  var BLOCK = 1000;
  for (var i = 0; i < values.length; i += BLOCK) {
    var part = values.slice(i, i + BLOCK);
    setValuesGrowing_(sh, 2 + i, part.length, headers.length, part);
  }
}

/**
 * Admin: moves the finished data dated before payload.cutoff to a new spreadsheet in the Drive folder "back up data", then
 * removes it from the main sheets. payload: {cutoff, username, password} — the username and password of the admin who is
 * logged in are asked again first. Returns the backup entry (name, sheet id, link, counts).
 */
function backupData(token, payload) {
  return safeCall_(function () {
    var session = requireAuth_(token, 'admin');
    payload = payload || {};
    requireDateStr_(payload.cutoff, 'วันที่ตัดยอด');
    var cutoff = payload.cutoff;
    if (cutoff > todayStr_()) throw ValidationError_('วันที่ตัดยอดต้องไม่เกินวันนี้');
    // confirm again: the logged-in admin's own username + password
    var user = findStaffRow_(getRows_(getSheet_(CONFIG.SHEETS.STAFF)), session.staffId, { username: session.username });
    if (String(payload.username || '').trim().toLowerCase() !== String(user.Username).toLowerCase()) throw ValidationError_('ชื่อผู้ใช้ไม่ตรงกับบัญชีที่เข้าสู่ระบบอยู่');
    if (hashPassword_(String(payload.password || ''), user.PasswordSalt) !== user.PasswordHash) throw ValidationError_('รหัสผ่านไม่ถูกต้อง');

    var lock = LockService.getScriptLock();
    lock.waitLock(30000);
    var file = null;
    try {
      var bsheet = getSheet_(CONFIG.SHEETS.BACKUPS);
      var pending = getRows_(bsheet).filter(function (b) { return b.Status !== 'เสร็จสิ้น'; })[0];
      if (pending) throw ValidationError_('มี backup ' + pending.FileName + ' ที่ยังทำไม่เสร็จ (' + pending.Status + ') — ตรวจสอบไฟล์นั้นก่อนทำ backup ใหม่');
      var plan = backupPlan_(cutoff);
      var total = 0;
      BACKUP_SHEETS_.forEach(function (b) { total += plan.move[b[0]].length; });
      if (!total) throw ValidationError_('ไม่มีข้อมูลที่จบแล้วก่อนวันที่ ' + cutoff + ' ให้ย้าย');

      var props = PropertiesService.getScriptProperties();
      var rootId = props.getProperty(CONFIG.PROPS.ROOT_FOLDER_ID);
      var root = null;
      if (rootId) { try { root = DriveApp.getFolderById(rootId); } catch (e) { root = null; } }
      if (!root) root = DriveApp.getFolderById(ensureDriveFolders_().rootId);
      var folder = getOrCreateChildFolder_(root, BACKUP_FOLDER_NAME_);

      // 1. the new spreadsheet, filled and checked BEFORE anything is removed
      var fileName = 'ข้อมูลตั้งแต่ ' + plan.from + ' - ' + plan.to;
      var ss = SpreadsheetApp.create(fileName);
      file = DriveApp.getFileById(ss.getId());
      file.moveTo(folder);
      var first = true;
      BACKUP_SHEETS_.forEach(function (b) {
        var headers = CONFIG.HEADERS[b[0]], rows = plan.move[b[0]];
        var sh = first ? ss.getSheets()[0].setName(b[0]) : ss.insertSheet(b[0]);
        first = false;
        sh.getRange(1, 1, 1, headers.length).setValues([headers]).setFontWeight('bold');
        sh.setFrozenRows(1);
        (CONFIG.TEXT_COLUMNS[b[0]] || []).forEach(function (col) {
          var idx = headers.indexOf(col);
          if (idx > -1 && rows.length) sh.getRange(2, idx + 1, rows.length, 1).setNumberFormat('@');
        });
        if (rows.length) backupWriteBlocks_(sh, headers, backupValues_(b[0], headers, rows));
      });
      SpreadsheetApp.flush();
      BACKUP_SHEETS_.forEach(function (b) {
        var written = ss.getSheetByName(b[0]).getLastRow() - 1;
        if (written !== plan.move[b[0]].length) throw new Error('ตรวจสอบไฟล์ backup ไม่ผ่าน: ' + b[0] + ' เขียนได้ ' + written + ' จาก ' + plan.move[b[0]].length + ' แถว (ยังไม่ได้ลบข้อมูลเดิม)');
      });

      // 2. what the dashboard needs, then the registry entry
      var backupId = 'BK' + nextSeq_('BACKUP', 4);
      var counts = backupCounts_(plan);
      var summary = backupSummaryRows_(plan, backupId);
      if (summary.length) appendRows_(getSheet_(CONFIG.SHEETS.DAILY_SUMMARY), CONFIG.HEADERS.DailySummary, summary);
      var entry = {
        BackupID: backupId, CreatedAt: nowStr_(), CreatedBy: session.username, FromDate: plan.from, ToDate: plan.to, CutoffDate: cutoff,
        FileName: fileName, SpreadsheetID: ss.getId(), SpreadsheetUrl: ss.getUrl(), FolderID: folder.getId(), CountsJSON: JSON.stringify(counts), Status: 'กำลังล้างข้อมูลเดิม'
      };
      var entryRow = appendRow_(bsheet, CONFIG.HEADERS.Backups, entry);

      // 3. remove what was moved: each sheet is rewritten with only the rows that stay
      BACKUP_SHEETS_.forEach(function (b) {
        var headers = CONFIG.HEADERS[b[0]], sh = getSheet_(CONFIG.SHEETS[b[1]]);
        if (!plan.move[b[0]].length) return;
        var keep = backupValues_(b[0], headers, plan.keep[b[0]]);
        var last = sh.getLastRow();
        if (last > 1) sh.getRange(2, 1, last - 1, headers.length).clearContent();
        if (keep.length) {
          resetNumberFormats_(sh, headers, 2, keep.length);
          backupWriteBlocks_(sh, headers, keep);
        }
      });
      entry.Status = 'เสร็จสิ้น';
      updateRowByIndex_(bsheet, CONFIG.HEADERS.Backups, entryRow, entry);
      writeAudit_(session, 'BACKUP', 'Backup', backupId, fileName + ' · ' + total + ' แถว');
      bumpDashboardCacheVersion_();
      invalidateStockLiteCache_();
      return { ok: true, data: mapBackupOut_(entry) };
    } finally {
      lock.releaseLock();
    }
  });
}
