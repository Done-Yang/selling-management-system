/**
 * 🩺 ລະບົບຕິດຕາມການດູແລລູກຄ້າ (Google Sheets + Apps Script)
 *
 * ວິທີຕິດຕັ້ງ: Extensions → Apps Script → ວາງໂຄດນີ້ → ເລືອກ ตั้งค่าระบบ → Run → ອະນຸຍາດ
 * ຫຼັງຈາກນັ້ນໃຊ້ເມນູ "🩺 ລະບົບດູແລ" ຫຼື ໝາຍ ☑ ໃນ sheet "ປ້ອນຂໍ້ມູນ"
 */

// ───────────────────────── ຄ່າຄົງທີ່ ─────────────────────────
const SH = {
  INPUT: 'ປ້ອນຂໍ້ມູນ',
  DATA: 'ຂໍ້ມູນລູກຄ້າ',
  REPORT: 'ລາຍງານ',
  SETTINGS: 'ຕັ້ງຄ່າ',
};

const DEFAULT_TYPES = ['ກົດໄຫຼຍ້ອນ', 'ກະເພາະ', 'ໄທລອຍ', 'ພູມແພ້', 'ສະເກັດເງິນ'];
const DEFAULT_SETS = [10, 20, 30, 60];
const CUST_TYPES = ['ລູກຄ້າໃໝ່', 'ລູກຄ້າເກົ່າ'];
const CARE_STATUS = ['ດູແລຕໍ່', 'ຢຸດກິນ'];
const SYMPTOM_STATUS = ['ຍັງບໍ່ດີຂຶ້ນ', 'ດີຂຶ້ນ', 'ດີແລ້ວ'];
const ALL = 'ທັງໝົດ';

// ຖັນໃນ sheet ຂໍ້ມູນລູກຄ້າ (A..O ປ້ອນໂດຍ script, P..T ເປັນສູດ)
const DATA_HEADERS = [
  'ID', 'ວັນບັນທຶກ', 'ເບີຕິດຕໍ່', 'ວັນທີເລີ່ມດູແລ', 'ປະເພດ', 'ເຊັດດູແລ (ມື້)',
  'ວັນຮັບສິນຄ້າ', 'ວັນສິນຄ້າໝົດ', 'ອາການລ່າສຸດ', 'ຕົ້ນທຶນ', 'ຍອດຂາຍ',
  'ລູກຄ້າເກົ່າ/ໃໝ່', 'ສະຖານະດູແລ', 'ວັນທີອັບເດດລ່າສຸດ', 'ສະຖານະອາການ',
];
const DATA_COLS = DATA_HEADERS.length; // 15

// ຕຳແໜ່ງຊ່ອງໃນ sheet ປ້ອນຂໍ້ມູນ (ຖັນ C)
const IN = {
  ID: 3, PHONE: 4, START: 5, TYPE: 6, SET: 7, RECV: 8, END: 9,
  SYMPTOM: 10, SYM_STATUS: 11, COST: 12, SALES: 13, CUST: 14, STATUS: 15, UPDATED: 16, PROFIT: 17,
  BTN_SAVE: 19, BTN_LOAD: 20, BTN_CLEAR: 21, MSG: 23,
};

const COLOR = {
  primary: '#2E7D5B',
  primaryLight: '#E6F2EC',
  accent: '#F2A541',
  grey: '#F1F3F4',
  border: '#C9D6CF',
  warn: '#FCE8D2',
  danger: '#F8D7D3',
};
const FONT = 'Noto Sans Lao';
const MONEY = '#,##0 "₭"';
const DATE_FMT = 'dd/MM/yyyy';
const DAYS_FMT = '0 "ມື້"';

const D = `'${SH.DATA}'!`;
const R = `'${SH.REPORT}'!`;

// ───────────────────────── ເມນູ / trigger ─────────────────────────
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('🩺 ລະບົບດູແລ')
    .addItem('💾 ບັນທຶກ', 'บันทึกข้อมูล')
    .addItem('🔍 ໂຫຼດມາແກ້ໄຂ (ID ຫຼື ເບີ)', 'โหลดมาแก้ไข')
    .addItem('🧹 ລ້າງຟອມ', 'ล้างฟอร์ม')
    .addSeparator()
    .addItem('⚙️ ຕັ້ງຄ່າລະບົບ (ສ້າງ/ສ້າງຄືນ sheet)', 'ตั้งค่าระบบ')
    .addToUi();
}

/** ປຸ່ມ ☑ ໃນ sheet ປ້ອນຂໍ້ມູນ + ກວດລູກຄ້າເກົ່າ/ໃໝ່ອັດຕະໂນມັດ */
function onEdit(e) {
  if (!e || !e.range) return;
  const sh = e.range.getSheet();
  if (sh.getName() !== SH.INPUT || e.range.getColumn() !== 3 || e.range.getNumRows() !== 1) return;
  const row = e.range.getRow();

  if (e.value === 'TRUE' && [IN.BTN_SAVE, IN.BTN_LOAD, IN.BTN_CLEAR].includes(row)) {
    e.range.setValue(false);
    if (row === IN.BTN_SAVE) บันทึกข้อมูล();
    if (row === IN.BTN_LOAD) โหลดมาแก้ไข();
    if (row === IN.BTN_CLEAR) ล้างฟอร์ม();
    return;
  }

  if (row === IN.PHONE && !sh.getRange(IN.ID, 3).getValue()) {
    const phone = normalizePhone_(e.range.getValue());
    if (!phone) return;
    const known = findRows_(r => r[2] === phone).length > 0;
    sh.getRange(IN.CUST, 3).setValue(known ? 'ລູກຄ້າເກົ່າ' : 'ລູກຄ້າໃໝ່');
  }
}

// ───────────────────────── ບັນທຶກ / ໂຫຼດ / ລ້າງ ─────────────────────────
function บันทึกข้อมูล() {
  const ss = SpreadsheetApp.getActive();
  const inp = ss.getSheetByName(SH.INPUT);
  const data = ss.getSheetByName(SH.DATA);
  if (!inp || !data) return notify_('ກະລຸນາ Run ⚙️ ຕັ້ງຄ່າລະບົບ ກ່ອນ', true);

  const v = inp.getRange(IN.ID, 3, IN.STATUS - IN.ID + 1, 1).getValues().map(r => r[0]);
  const at = row => v[row - IN.ID];

  const id = String(at(IN.ID) || '').trim();
  const phone = normalizePhone_(at(IN.PHONE));
  const start = at(IN.START);
  const type = at(IN.TYPE);
  const setDays = Number(at(IN.SET));
  const recv = at(IN.RECV);
  const symptom = at(IN.SYMPTOM);
  const symStatus = at(IN.SYM_STATUS);
  const cost = at(IN.COST) === '' ? 0 : at(IN.COST);
  const sales = at(IN.SALES) === '' ? 0 : at(IN.SALES);
  let cust = at(IN.CUST);
  const status = at(IN.STATUS) || 'ດູແລຕໍ່';

  const errors = [];
  if (!phone) errors.push('ເບີຕິດຕໍ່');
  if (!(start instanceof Date)) errors.push('ວັນທີເລີ່ມດູແລ (ຕ້ອງເປັນວັນທີ)');
  if (!type) errors.push('ປະເພດ');
  if (!setDays) errors.push('ເຊັດດູແລ');
  if (!(recv instanceof Date)) errors.push('ວັນຮັບສິນຄ້າ (ຕ້ອງເປັນວັນທີ)');
  if (typeof cost !== 'number') errors.push('ຕົ້ນທຶນ (ຕ້ອງເປັນຕົວເລກ)');
  if (typeof sales !== 'number') errors.push('ຍອດຂາຍ (ຕ້ອງເປັນຕົວເລກ)');
  if (errors.length) return notify_('❌ ກະລຸນາກວດ: ' + errors.join(', '), true);

  const end = new Date(recv.getTime() + setDays * 86400000);

  const lock = LockService.getDocumentLock();
  lock.waitLock(15000);
  try {
    const ids = getIdColumn_(data);
    if (!cust) {
      const known = findRows_(r => r[2] === phone && r[0] !== id).length > 0;
      cust = known ? 'ລູກຄ້າເກົ່າ' : 'ລູກຄ້າໃໝ່';
    }

    let row, recId, createdAt = new Date();
    if (id) {
      const idx = ids.indexOf(id);
      if (idx < 0) return notify_(`❌ ບໍ່ພົບ ID ${id} (ລ້າງຊ່ອງ ID ຖ້າຕ້ອງການບັນທຶກເປັນລາຍການໃໝ່)`, true);
      row = idx + 2;
      recId = id;
      createdAt = data.getRange(row, 2).getValue() || createdAt;
    } else {
      row = lastIdRow_(ids) + 1;
      recId = nextId_(ids);
      if (row > data.getMaxRows()) data.insertRowsAfter(data.getMaxRows(), 100);
    }

    data.getRange(row, 1, 1, DATA_COLS).setValues([[
      recId, createdAt, phone, start, type, setDays, recv, end, symptom, cost, sales, cust, status, new Date(), symStatus,
    ]]);
    SpreadsheetApp.flush();
  } finally {
    lock.releaseLock();
  }

  ล้างฟอร์ม(true);
  notify_(id ? `✅ ອັບເດດ ${id} ສຳເລັດ` : `✅ ບັນທຶກສຳເລັດ (${phone})`);
}

/** ໂຫຼດລາຍການມາແກ້ໄຂ: ໃຊ້ ID ໃນ C3, ຖ້າບໍ່ມີ ໃຊ້ເບີໃນ C4 (ເອົາລາຍການລ່າສຸດ) */
function โหลดมาแก้ไข() {
  const ss = SpreadsheetApp.getActive();
  const inp = ss.getSheetByName(SH.INPUT);
  const id = String(inp.getRange(IN.ID, 3).getValue() || '').trim();
  const phone = normalizePhone_(inp.getRange(IN.PHONE, 3).getValue());
  if (!id && !phone) return notify_('❌ ໃສ່ ID (C3) ຫຼື ເບີຕິດຕໍ່ (C4) ກ່ອນ ແລ້ວກົດໂຫຼດ', true);

  const matches = id ? findRows_(r => r[0] === id) : findRows_(r => r[2] === phone);
  if (!matches.length) return notify_(`❌ ບໍ່ພົບ ${id || phone}`, true);
  const r = matches[matches.length - 1].values;

  inp.getRange(IN.ID, 3, 6, 1).setValues([[r[0]], [r[2]], [r[3]], [r[4]], [r[5]], [r[6]]]);
  inp.getRange(IN.SYMPTOM, 3, 7, 1).setValues([[r[8]], [r[14]], [r[9]], [r[10]], [r[11]], [r[12]], [r[13]]]);
  notify_(`🔍 ໂຫຼດ ${r[0]} ແລ້ວ — ແກ້ໄຂ ແລ້ວກົດ 💾 ບັນທຶກ ເພື່ອອັບເດດ`);
}

function ล้างฟอร์ม(silent) {
  const ss = SpreadsheetApp.getActive();
  const inp = ss.getSheetByName(SH.INPUT);
  const today = today_();
  inp.getRange(IN.ID, 3, IN.RECV - IN.ID + 1, 1).clearContent();
  inp.getRange(IN.SYMPTOM, 3, IN.UPDATED - IN.SYMPTOM + 1, 1).clearContent();
  inp.getRange(IN.START, 3).setValue(today);
  inp.getRange(IN.RECV, 3).setValue(today);
  inp.getRange(IN.STATUS, 3).setValue('ດູແລຕໍ່');
  inp.getRange(IN.PHONE, 3).activate();
  if (silent !== true) notify_('🧹 ລ້າງຟອມແລ້ວ');
}

// ───────────────────────── ຕິດຕັ້ງລະບົບ ─────────────────────────
function ตั้งค่าระบบ() {
  const ss = SpreadsheetApp.getActive();
  ss.setSpreadsheetLocale('en_GB'); // ພິມວັນທີແບບ ວັນ/ເດືອນ/ປີ ໄດ້ ແລະ ສູດເຮັດວຽກຄືກັນທຸກເຄື່ອງ

  const isNewSettings = !ss.getSheetByName(SH.SETTINGS);
  const sInput = getOrCreate_(ss, SH.INPUT);
  const sData = getOrCreate_(ss, SH.DATA);
  const sReport = getOrCreate_(ss, SH.REPORT);
  const sSet = getOrCreate_(ss, SH.SETTINGS);

  buildSettings_(sSet, isNewSettings);
  buildData_(sData);
  buildReport_(sReport, sSet);
  buildInput_(sInput, sSet);

  // ລຽງລຳດັບ sheet ແລະ ລົບ sheet ເປົ່າເລີ່ມຕົ້ນ
  [sInput, sData, sReport, sSet].forEach((s, i) => { ss.setActiveSheet(s); ss.moveActiveSheet(i + 1); });
  ss.getSheets().forEach(s => {
    const ours = Object.values(SH).includes(s.getName());
    if (!ours && s.getLastRow() === 0 && s.getLastColumn() === 0 && ss.getSheets().length > 4) ss.deleteSheet(s);
  });

  ss.setActiveSheet(sInput);
  try { onOpen(); } catch (err) { /* Run ຈາກ editor: ເມນູຈະຂຶ້ນເມື່ອເປີດ sheet ຄືນ */ }
  ล้างฟอร์ม(true);
  notify_('✅ ຕັ້ງຄ່າລະບົບສຳເລັດ! ເລີ່ມປ້ອນຂໍ້ມູນໄດ້ເລີຍ');
}

function buildSettings_(sh, isNew) {
  if (isNew) {
    sh.clear();
    sh.getRange('A1').setValue('ປະເພດ (ອາການ)');
    sh.getRange(2, 1, DEFAULT_TYPES.length, 1).setValues(DEFAULT_TYPES.map(t => [t]));
    sh.getRange('C1').setValue('ເຊັດດູແລ (ມື້)');
    sh.getRange(2, 3, DEFAULT_SETS.length, 1).setValues(DEFAULT_SETS.map(n => [n]));
    sh.getRange('I1').setValue('💡 ເພີ່ມ/ແກ້ ປະເພດ ໃນຖັນ A ແລະ ເຊັດ (ຕົວເລກມື້) ໃນຖັນ C — dropdown ຈະອັບເດດເອງ');
  }
  // ລາຍການສຳລັບ filter ໃນລາຍງານ (ສູດ — ບໍ່ຕ້ອງແກ້)
  sh.getRange('E1:E2').setValues([['ລູກຄ້າ (filter)'], [ALL]]);
  sh.getRange('E3').setFormula(`=IFERROR(SORT(UNIQUE(FILTER(${D}C2:C, ${D}C2:C<>""))),"")`);
  sh.getRange('G1:G2').setValues([['ເຊັດ (filter)'], [ALL]]);
  sh.getRange('G3').setFormula('=IFERROR(FILTER(C2:C, C2:C<>""),"")');

  sh.getRange('E3:E').setNumberFormat('@');
  sh.getRange('E2:G').setFontColor('#888888');
  ['A1', 'C1', 'E1', 'G1'].forEach(a => sh.getRange(a)
    .setFontWeight('bold').setFontColor('#FFFFFF').setBackground(COLOR.primary));
  sh.getRange('I1').setFontColor('#555555');
  sh.getRange('A:I').setFontFamily(FONT);
  sh.setColumnWidths(1, 7, 140);
  sh.setFrozenRows(1);
}

function buildData_(sh) {
  sh.getRange(1, 1, 1, DATA_COLS).setValues([DATA_HEADERS]);
  migrateCareStatus_(sh);
  const f = (col, header, expr) => sh.getRange(col + '1').setFormula(`={"${header}"; ARRAYFORMULA(IF(A2:A="","",${expr}))}`);
  f('P', 'ກຳໄລ', 'K2:K-J2:J');
  f('Q', 'ມື້ທີ່ເຫຼືອ', 'IF(M2:M<>"ດູແລຕໍ່","",H2:H-TODAY())');
  f('R', 'ເດືອນ', 'TEXT(D2:D,"yyyy-mm")');
  f('S', 'ສະຖານະ', 'IF(M2:M="ຢຸດກິນ","ສິ້ນສຸດດູແລ","ກຳລັງດູແລ")');
  // ຖັນ T: 1 = ຜ່ານ filter ໃນ sheet ລາຍງານ (ຖັນຊ່ວຍ, ເຊື່ອງໄວ້)
  sh.getRange('T1').setFormula(
    `={"_filter"; ARRAYFORMULA(IF(A2:A="",0,` +
    `(D2:D>=IF(${R}$B$4="",DATE(1900,1,1),${R}$B$4))*` +
    `(D2:D<=IF(${R}$C$4="",DATE(9999,12,31),${R}$C$4))*` +
    `((${R}$D$4="")+(${R}$D$4="${ALL}")+(C2:C&""=${R}$D$4&"")>0)*` +
    `((${R}$E$4="")+(${R}$E$4="${ALL}")+(F2:F&""=${R}$E$4&"")>0)))}`
  );

  // ຮູບແບບ
  sh.getRange('A1:T1').setFontWeight('bold').setFontColor('#FFFFFF').setBackground(COLOR.primary)
    .setVerticalAlignment('middle').setWrap(true);
  sh.getRange('P1:S1').setBackground('#3F9B72');
  sh.setRowHeight(1, 36);
  sh.setFrozenRows(1);
  sh.setFrozenColumns(3);
  sh.getRange('A:T').setFontFamily(FONT);
  ['B2:B', 'N2:N'].forEach(a => sh.getRange(a).setNumberFormat('dd/MM/yyyy HH:mm'));
  sh.getRange('C2:C').setNumberFormat('@');
  ['D2:D', 'G2:G', 'H2:H'].forEach(a => sh.getRange(a).setNumberFormat(DATE_FMT));
  sh.getRange('F2:F').setNumberFormat(DAYS_FMT);
  ['J2:J', 'K2:K', 'P2:P'].forEach(a => sh.getRange(a).setNumberFormat(MONEY));
  sh.getRange('O2:O').setNumberFormat('General');
  sh.getRange('Q2:Q').setNumberFormat('0');
  sh.getRange('I2:I').setWrap(true);
  sh.showColumns(1, 20); // ສຳລັບລະບົບເກົ່າທີ່ເຄີຍເຊື່ອງຖັນອື່ນ
  sh.setColumnWidths(1, 20, 110);
  sh.setColumnWidth(9, 220);
  sh.setColumnWidth(14, 130);
  sh.hideColumns(20);

  // ສີແຖວ: ໝົດແລ້ວ (ແດງ), ໃກ້ໝົດ ≤ 3 ມື້ (ສົ້ມ), ອາການດີແລ້ວ (ຂຽວ), ຢຸດກິນ (ເທົາ)
  const rng = sh.getRange('A2:S');
  sh.setConditionalFormatRules([
    SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied('=AND($A2<>"",$M2="ດູແລຕໍ່",$Q2<>"",$Q2<0)')
      .setBackground(COLOR.danger).setRanges([rng]).build(),
    SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied('=AND($A2<>"",$M2="ດູແລຕໍ່",$Q2<>"",$Q2<=3)')
      .setBackground(COLOR.warn).setRanges([rng]).build(),
    SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied('=AND($A2<>"",$O2="ດີແລ້ວ")')
      .setFontColor(COLOR.primary).setRanges([rng]).build(),
    SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied('=AND($A2<>"",$M2="ຢຸດກິນ")')
      .setFontColor('#9AA0A6').setRanges([rng]).build(),
  ]);

  sh.getProtections(SpreadsheetApp.ProtectionType.SHEET).forEach(p => p.remove());
  sh.protect().setDescription('ແກ້ໄຂຜ່ານ sheet ປ້ອນຂໍ້ມູນ').setWarningOnly(true);
}

function buildInput_(sh, sSet) {
  sh.getProtections(SpreadsheetApp.ProtectionType.RANGE).forEach(p => p.remove());
  sh.clear();
  sh.clearConditionalFormatRules();
  sh.getRange('A:E').clearDataValidations();

  sh.setColumnWidth(1, 24);
  sh.setColumnWidth(2, 190);
  sh.setColumnWidth(3, 260);
  sh.setColumnWidth(4, 300);
  sh.getRange('A:E').setFontFamily(FONT).setVerticalAlignment('middle');

  sh.getRange('B1:D1').merge().setValue('🩺 ຟອມປ້ອນຂໍ້ມູນການດູແລລູກຄ້າ')
    .setFontSize(16).setFontWeight('bold').setFontColor('#FFFFFF').setBackground(COLOR.primary);
  sh.setRowHeight(1, 44);
  sh.getRange('B2:D2').merge().setValue('ປ້ອນຂໍ້ມູນໃນຖັນ C (ສີຂາວ) ແລ້ວໝາຍ ☑ "ບັນທຶກ" ທາງລຸ່ມ  •  * = ຈຳເປັນ')
    .setFontColor('#555555').setBackground(COLOR.primaryLight);

  const rows = [
    [IN.ID, 'ID (ອັດຕະໂນມັດ)', 'ວ່າງ = ລາຍການໃໝ່ • ມີ ID = ອັບເດດລາຍການເດີມ'],
    [IN.PHONE, 'ເບີຕິດຕໍ່ *', 'ລະບົບກວດລູກຄ້າເກົ່າ/ໃໝ່ໃຫ້ອັດຕະໂນມັດ'],
    [IN.START, 'ວັນທີເລີ່ມດູແລ *', 'ຄລິກສອງເທື່ອເພື່ອເລືອກວັນທີ'],
    [IN.TYPE, 'ປະເພດ *', 'ເພີ່ມປະເພດໄດ້ທີ່ sheet ຕັ້ງຄ່າ'],
    [IN.SET, 'ເຊັດດູແລ *', '10 / 20 / 30 / 60 ມື້'],
    [IN.RECV, 'ວັນຮັບສິນຄ້າ *', ''],
    [IN.END, 'ວັນສິນຄ້າໝົດ', '🔒 ຄຳນວນອັດຕະໂນມັດ = ວັນຮັບສິນຄ້າ + ເຊັດດູແລ'],
    [IN.SYMPTOM, 'ອາການລ່າສຸດ', ''],
    [IN.SYM_STATUS, 'ສະຖານະອາການ', 'ຍັງບໍ່ດີຂຶ້ນ / ດີຂຶ້ນ / ດີແລ້ວ'],
    [IN.COST, 'ຕົ້ນທຶນ', 'ກີບ'],
    [IN.SALES, 'ຍອດຂາຍຂອງລາຍການ', 'ກີບ'],
    [IN.CUST, 'ລູກຄ້າເກົ່າ/ໃໝ່', 'ເລືອກເອງ ຫຼື ປ່ອຍໃຫ້ລະບົບກວດ'],
    [IN.STATUS, 'ສະຖານະດູແລ', 'ດູແລຕໍ່ / ຢຸດກິນ'],
    [IN.UPDATED, 'ວັນທີອັບເດດລ່າສຸດ', '🔒 ບັນທຶກອັດຕະໂນມັດທຸກຄັ້ງທີ່ກົດ 💾 ບັນທຶກ'],
    [IN.PROFIT, 'ກຳໄລ', '🔒 ຄຳນວນອັດຕະໂນມັດ = ຍອດຂາຍ − ຕົ້ນທຶນ'],
  ];
  rows.forEach(([r, label, hint]) => {
    sh.getRange(r, 2).setValue(label).setFontWeight('bold').setBackground(COLOR.primaryLight);
    sh.getRange(r, 4).setValue(hint).setFontColor('#888888').setFontSize(9);
    sh.setRowHeight(r, 32);
  });
  sh.getRange(IN.ID, 2, IN.PROFIT - IN.ID + 1, 3)
    .setBorder(true, true, true, true, true, true, COLOR.border, SpreadsheetApp.BorderStyle.SOLID);

  const c = r => sh.getRange(r, 3);
  const dv = () => SpreadsheetApp.newDataValidation().setAllowInvalid(false);

  c(IN.ID).setBackground(COLOR.grey).setFontColor('#555555');
  c(IN.PHONE).setNumberFormat('@');
  [IN.START, IN.RECV].forEach(r => c(r).setNumberFormat(DATE_FMT)
    .setDataValidation(dv().requireDate().setHelpText('ໃສ່ວັນທີ ເຊັ່ນ 18/10/2026').build()));
  c(IN.TYPE).setDataValidation(dv().requireValueInRange(sSet.getRange('A2:A100'), true).build());
  c(IN.SET).setNumberFormat(DAYS_FMT)
    .setDataValidation(dv().requireValueInRange(sSet.getRange('C2:C30'), true).build());
  c(IN.END).setFormula(`=IF(AND(ISDATE(C${IN.RECV}),ISNUMBER(C${IN.SET})),C${IN.RECV}+C${IN.SET},"")`)
    .setNumberFormat(DATE_FMT).setBackground(COLOR.grey).setFontWeight('bold').setFontColor(COLOR.primary);
  c(IN.SYMPTOM).setWrap(true);
  sh.setRowHeight(IN.SYMPTOM, 56);
  [IN.COST, IN.SALES].forEach(r => c(r).setNumberFormat(MONEY)
    .setDataValidation(dv().requireNumberGreaterThanOrEqualTo(0).setHelpText('ໃສ່ຕົວເລກ').build()));
  c(IN.CUST).setDataValidation(dv().requireValueInList(CUST_TYPES, true).build());
  c(IN.SYM_STATUS).setDataValidation(dv().requireValueInList(SYMPTOM_STATUS, true).build());
  c(IN.STATUS).setDataValidation(dv().requireValueInList(CARE_STATUS, true).build());
  c(IN.UPDATED).setNumberFormat('dd/MM/yyyy HH:mm').setBackground(COLOR.grey).setFontColor('#555555');
  c(IN.PROFIT).setFormula(`=IF(OR(C${IN.SALES}<>"",C${IN.COST}<>""),N(C${IN.SALES})-N(C${IN.COST}),"")`)
    .setNumberFormat(MONEY).setBackground(COLOR.grey).setFontWeight('bold').setFontColor(COLOR.primary);

  // ລັອກຊ່ອງສູດ
  [IN.END, IN.PROFIT].forEach(r => {
    const p = c(r).protect().setDescription('ຄຳນວນອັດຕະໂນມັດ');
    p.addEditor(Session.getEffectiveUser());
    p.removeEditors(p.getEditors());
    if (p.canDomainEdit()) p.setDomainEdit(false);
  });

  // ປຸ່ມ ☑
  const buttons = [
    [IN.BTN_SAVE, '💾 ບັນທຶກ', '#D7F0E2', 'ໝາຍ ☑ ເພື່ອບັນທຶກ / ອັບເດດ'],
    [IN.BTN_LOAD, '🔍 ໂຫຼດມາແກ້ໄຂ', '#FFF4DE', 'ໃສ່ ID (C3) ຫຼື ເບີ (C4) ແລ້ວໝາຍ ☑ → ໂຫຼດລາຍການລ່າສຸດ'],
    [IN.BTN_CLEAR, '🧹 ລ້າງຟອມ', COLOR.grey, 'ໝາຍ ☑ ເພື່ອລ້າງທຸກຊ່ອງ'],
  ];
  buttons.forEach(([r, label, bg, hint]) => {
    sh.getRange(r, 2).setValue(label).setFontWeight('bold').setFontSize(12).setBackground(bg);
    sh.getRange(r, 3).insertCheckboxes().setValue(false).setBackground(bg).setHorizontalAlignment('left');
    sh.getRange(r, 4).setValue(hint).setFontColor('#888888').setFontSize(9);
    sh.setRowHeight(r, 34);
  });
  sh.getRange(IN.MSG, 2, 1, 3).merge().setFontWeight('bold').setFontSize(11);

  sh.setFrozenRows(2);
  sh.setHiddenGridlines(true);
}

function buildReport_(sh, sSet) {
  sh.getCharts().forEach(ch => sh.removeChart(ch));
  sh.clear();
  sh.clearConditionalFormatRules();
  sh.getRange('A:P').clearDataValidations();
  sh.setColumnWidth(1, 24);
  sh.setColumnWidths(2, 15, 125);
  sh.setColumnWidth(9, 200);
  sh.getRange('A:P').setFontFamily(FONT).setVerticalAlignment('middle');

  sh.getRange('B1:P1').merge().setValue('📊 ລາຍງານສະຫຼຸບການດູແລລູກຄ້າ')
    .setFontSize(16).setFontWeight('bold').setFontColor('#FFFFFF').setBackground(COLOR.primary);
  sh.setRowHeight(1, 44);

  // ── Filter ──
  sh.getRange('B3:E3').setValues([['📅 ຈາກວັນທີ', '📅 ຫາວັນທີ', '👤 ລູກຄ້າ (ເບີ)', '📦 ເຊັດດູແລ']])
    .setFontWeight('bold').setBackground(COLOR.primaryLight);
  const y = new Date().getFullYear();
  const tz = SpreadsheetApp.getActive().getSpreadsheetTimeZone();
  sh.getRange('B4:E4').setValues([[
    Utilities.parseDate(`${y}-01-01`, tz, 'yyyy-MM-dd'),
    Utilities.parseDate(`${y}-12-31`, tz, 'yyyy-MM-dd'),
    ALL, ALL,
  ]]).setBackground('#FFFDF5').setFontWeight('bold');
  const dv = () => SpreadsheetApp.newDataValidation().setAllowInvalid(false);
  sh.getRange('B4:C4').setNumberFormat(DATE_FMT).setDataValidation(dv().requireDate().build());
  sh.getRange('D4').setNumberFormat('@').setDataValidation(dv().requireValueInRange(sSet.getRange('E2:E1000'), true).build());
  sh.getRange('E4').setNumberFormat('0').setDataValidation(dv().requireValueInRange(sSet.getRange('G2:G50'), true).build());
  sh.getRange('F4').setValue('← ປ່ຽນ filter ໄດ້ (ລຶບວັນທີ = ບໍ່ຈຳກັດ) • ກອງຕາມ ວັນທີເລີ່ມດູແລ')
    .setFontColor('#888888').setFontSize(9);
  sh.getRange('B3:E4').setBorder(true, true, true, true, true, true, COLOR.border, SpreadsheetApp.BorderStyle.SOLID);

  // ── KPI ──
  // ຖັນ sheet ຂໍ້ມູນ: C = ເບີ, J = ຕົ້ນທຶນ, K = ຍອດຂາຍ, O = ສະຖານະອາການ,
  //                   R = ເດືອນ, S = ສະຖານະດູແລ, T = ຜ່ານ filter
  // ສະຖານະ ນັບ 1 ຄົນ ຕາມລາຍການລ່າສຸດຂອງເບີນັ້ນ (ວັນທີເລີ່ມດູແລ ຫຼ້າສຸດ) → ຄົນດຽວບໍ່ຖືກນັບຊ້ຳ
  const uniq = cond => `IFERROR(COUNTUNIQUE(FILTER(${D}C2:C, ${D}T2:T=1${cond})),0)`;
  const latest = cond =>
    `SORTN(SORT(FILTER({${D}C2:C, ${D}D2:D+ROW(${D}D2:D)/1000000, ${D}S2:S, ${D}O2:O}, ${D}T2:T=1${cond}), 2, FALSE), 9^9, 2, 1, TRUE)`;
  const CARE_COL = 3, SYM_COL = 4;
  const countLatest = (cond, col, val) => `IFERROR(ARRAYFORMULA(COUNTIF(INDEX(${latest(cond)}, 0, ${col}), "${val}")), 0)`;
  const kpis = [
    ['💰 ຕົ້ນທຶນ', `=SUMIFS(${D}J2:J, ${D}T2:T, 1)`, MONEY],
    ['💵 ຍອດຂາຍທັງໝົດ', `=SUMIFS(${D}K2:K, ${D}T2:T, 1)`, MONEY],
    ['📈 ກຳໄລ', `=SUMIFS(${D}K2:K, ${D}T2:T, 1)-SUMIFS(${D}J2:J, ${D}T2:T, 1)`, MONEY],
    ['👥 ລູກຄ້າທັງໝົດ', '=' + uniq(''), '#,##0 "ຄົນ"'],
    ['✅ ສິ້ນສຸດດູແລ', '=' + countLatest('', CARE_COL, 'ສິ້ນສຸດດູແລ'), '#,##0 "ຄົນ"'],
    ['🩺 ກຳລັງດູແລ', '=' + countLatest('', CARE_COL, 'ກຳລັງດູແລ'), '#,##0 "ຄົນ"'],
    ['😊 ອາການດີແລ້ວ', '=' + countLatest('', SYM_COL, 'ດີແລ້ວ'), '#,##0 "ຄົນ"'],
  ];
  kpis.forEach(([label, formula, fmt], i) => {
    sh.getRange(6, 2 + i).setValue(label);
    sh.getRange(7, 2 + i).setFormula(formula).setNumberFormat(fmt);
  });
  sh.getRange('B6:H6').setFontColor('#FFFFFF').setBackground('#3F9B72').setFontWeight('bold').setHorizontalAlignment('center');
  sh.getRange('B7:H7').setFontSize(15).setFontWeight('bold').setFontColor(COLOR.primary)
    .setBackground(COLOR.primaryLight).setHorizontalAlignment('center');
  sh.getRange('D7').setFontColor('#B26A00');
  sh.setRowHeight(7, 46);
  sh.getRange('B6:H7').setBorder(true, true, true, true, true, false, '#FFFFFF', SpreadsheetApp.BorderStyle.SOLID_THICK);

  // ── ສະຫຼຸບລາຍເດືອນ ──
  sh.getRange('B9').setValue('📆 ສະຫຼຸບລາຍເດືອນ').setFontWeight('bold').setFontSize(12);
  sh.getRange('B11:I11').setValues([['ເດືອນ', 'ຕົ້ນທຶນ', 'ຍອດຂາຍ', 'ກຳໄລ', 'ລູກຄ້າ', 'ສິ້ນສຸດດູແລ', 'ກຳລັງດູແລ', 'ອາການດີແລ້ວ']])
    .setFontWeight('bold').setFontColor('#FFFFFF').setBackground(COLOR.primary).setHorizontalAlignment('center');
  sh.getRange('B12').setFormula(
    `=IFERROR(QUERY({${D}R2:R, ${D}J2:J, ${D}K2:K, ${D}T2:T}, ` +
    `"select Col1, sum(Col2), sum(Col3), sum(Col3)-sum(Col2) where Col4 = 1 group by Col1 order by Col1 ` +
    `label sum(Col2) '', sum(Col3) '', sum(Col3)-sum(Col2) ''", 0), "")`
  );
  // ລາຍເດືອນ: ໃຊ້ລາຍການລ່າສຸດຂອງແຕ່ລະຄົນ ພາຍໃນເດືອນນັ້ນ
  const inMonth = `, ${D}R2:R=m`;
  const perMonth = expr => `=MAP(B12:B47, LAMBDA(m, IF(m="", "", ${expr})))`;
  sh.getRange('F12').setFormula(perMonth(uniq(inMonth)));
  sh.getRange('G12').setFormula(perMonth(countLatest(inMonth, CARE_COL, 'ສິ້ນສຸດດູແລ')));
  sh.getRange('H12').setFormula(perMonth(countLatest(inMonth, CARE_COL, 'ກຳລັງດູແລ')));
  sh.getRange('I12').setFormula(perMonth(countLatest(inMonth, SYM_COL, 'ດີແລ້ວ')));
  sh.getRange('C12:E47').setNumberFormat(MONEY);
  sh.getRange('F12:I47').setNumberFormat('0');
  sh.getRange('B12:B47').setHorizontalAlignment('center');
  sh.getRange('F12:I47').setHorizontalAlignment('center');
  sh.getRange('B11:I47').setBorder(true, true, true, true, null, true, COLOR.border, SpreadsheetApp.BorderStyle.SOLID);
  sh.setConditionalFormatRules([
    SpreadsheetApp.newConditionalFormatRule().whenFormulaSatisfied('=AND($B12<>"",ISEVEN(ROW()))')
      .setBackground('#F7FAF8').setRanges([sh.getRange('B12:I47')]).build(),
    SpreadsheetApp.newConditionalFormatRule().whenFormulaSatisfied('=AND($B51<>"",ISEVEN(ROW()))')
      .setBackground('#F7FAF8').setRanges([sh.getRange('B51:P')]).build(),
  ]);

  // ── Chart ຍອດຂາຍ vs ກຳໄລ ──
  const chart = sh.newChart()
    .setChartType(Charts.ChartType.COLUMN)
    .addRange(sh.getRange('B11:B47'))
    .addRange(sh.getRange('D11:E47'))
    .setMergeStrategy(Charts.ChartMergeStrategy.MERGE_COLUMNS)
    .setNumHeaders(1)
    .setPosition(9, 11, 0, 0)
    .setOption('title', 'ຍອດຂາຍ vs ກຳໄລ ລາຍເດືອນ')
    .setOption('colors', [COLOR.primary, COLOR.accent])
    .setOption('legend', { position: 'top' })
    .setOption('width', 560)
    .setOption('height', 320)
    .build();
  sh.insertChart(chart);

  // ── ລາຍການລະອຽດ (ຕາມ filter) ──
  sh.getRange('B49').setValue('📋 ລາຍການລະອຽດ (ຕາມ filter)').setFontWeight('bold').setFontSize(12);
  sh.getRange('B50:P50').setValues([[
    'ID', 'ເບີຕິດຕໍ່', 'ວັນທີເລີ່ມດູແລ', 'ປະເພດ', 'ເຊັດ', 'ວັນຮັບສິນຄ້າ', 'ວັນສິນຄ້າໝົດ',
    'ອາການລ່າສຸດ', 'ສະຖານະອາການ', 'ຕົ້ນທຶນ', 'ຍອດຂາຍ', 'ກຳໄລ', 'ເກົ່າ/ໃໝ່', 'ສະຖານະດູແລ', 'ອັບເດດລ່າສຸດ',
  ]]).setFontWeight('bold').setFontColor('#FFFFFF').setBackground(COLOR.primary).setHorizontalAlignment('center');
  sh.getRange('B51').setFormula(
    `=IFERROR(SORT(FILTER({${D}A2:A, ${D}C2:C, ${D}D2:D, ${D}E2:E, ${D}F2:F, ${D}G2:G, ${D}H2:H, ` +
    `${D}I2:I, ${D}O2:O, ${D}J2:J, ${D}K2:K, ${D}P2:P, ${D}L2:L, ${D}S2:S, ${D}N2:N}, ${D}T2:T=1), 3, FALSE), ` +
    `"ບໍ່ມີຂໍ້ມູນຕາມ filter")`
  );
  ['D51:D', 'G51:H'].forEach(a => sh.getRange(a).setNumberFormat(DATE_FMT));
  sh.getRange('F51:F').setNumberFormat(DAYS_FMT);
  sh.getRange('K51:M').setNumberFormat(MONEY);
  sh.getRange('P51:P').setNumberFormat('dd/MM/yyyy HH:mm');
  sh.getRange('C51:C').setNumberFormat('@');

  sh.setFrozenRows(4);
  sh.setHiddenGridlines(true);
}

// ───────────────────────── ຕົວຊ່ວຍ ─────────────────────────
/** ລະບົບເກົ່າເກັບ "ດີແລ້ວ" ໃນສະຖານະດູແລ → ຍ້າຍໄປສະຖານະອາການ ແລະ ຕັ້ງສະຖານະດູແລເປັນ "ຢຸດກິນ" */
function migrateCareStatus_(sh) {
  const last = lastIdRow_(getIdColumn_(sh));
  if (last < 2) return;
  const rng = sh.getRange(2, 13, last - 1, 3); // M..O
  const rows = rng.getValues();
  let changed = false;
  rows.forEach(r => {
    if (r[0] === 'ດີແລ້ວ') { r[0] = 'ຢຸດກິນ'; r[2] = 'ດີແລ້ວ'; changed = true; }
  });
  if (changed) rng.setValues(rows);
}

function getOrCreate_(ss, name) {
  return ss.getSheetByName(name) || ss.insertSheet(name);
}

function normalizePhone_(v) {
  return String(v == null ? '' : v).replace(/[\s\-.]/g, '').trim();
}

function today_() {
  const tz = SpreadsheetApp.getActive().getSpreadsheetTimeZone();
  return Utilities.parseDate(Utilities.formatDate(new Date(), tz, 'yyyy-MM-dd'), tz, 'yyyy-MM-dd');
}

function getIdColumn_(data) {
  const n = data.getMaxRows() - 1;
  return n > 0 ? data.getRange(2, 1, n, 1).getValues().map(r => String(r[0] || '')) : [];
}

function lastIdRow_(ids) {
  for (let i = ids.length - 1; i >= 0; i--) if (ids[i]) return i + 2;
  return 1;
}

function nextId_(ids) {
  const max = ids.reduce((m, id) => {
    const k = /^CS-(\d+)$/.exec(id);
    return k ? Math.max(m, Number(k[1])) : m;
  }, 0);
  return 'CS-' + String(max + 1).padStart(4, '0');
}

/** ຄືນຄ່າ [{row, values}] ຂອງແຖວໃນ sheet ຂໍ້ມູນທີ່ກົງເງື່ອນໄຂ (values: A..O, ເບີເປັນ string) */
function findRows_(pred) {
  const data = SpreadsheetApp.getActive().getSheetByName(SH.DATA);
  const last = lastIdRow_(getIdColumn_(data));
  if (last < 2) return [];
  return data.getRange(2, 1, last - 1, DATA_COLS).getValues()
    .map((values, i) => {
      values[0] = String(values[0] || '');
      values[2] = normalizePhone_(values[2]);
      return { row: i + 2, values };
    })
    .filter(r => r.values[0] && pred(r.values));
}

function notify_(msg, isError) {
  const ss = SpreadsheetApp.getActive();
  ss.toast(msg, isError ? '⚠️ ແຈ້ງເຕືອນ' : '🩺 ລະບົບດູແລ', 6);
  const inp = ss.getSheetByName(SH.INPUT);
  if (inp) {
    inp.getRange(IN.MSG, 2).setValue(msg)
      .setFontColor(isError ? '#C5221F' : COLOR.primary)
      .setBackground(isError ? COLOR.danger : COLOR.primaryLight);
  }
}
