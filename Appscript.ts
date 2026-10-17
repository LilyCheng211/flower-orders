// ============================================================
// 花店訂單管理系統 - Apps Script 完整版 v2
// 欄位：A姓名/B電話/C金額/D取花日期/E幾點取花/F配送/G品項/H LINE Pay 交易代碼/I收到款/J收花資訊/K卡片內容/L備註
// H 欄在前端 JSON 仍叫 last5。電話與交易代碼要以文字寫入，避免開頭的 0 被吃掉。
// M 欄「訂單編號」是日曆同步自己寫的，請不要改、也不要拿來填備註。刪掉一列時列號會往上移，所以不能用列號認訂單。
// 這份檔案要貼回 Apps Script，並建立新的部署版本，才會在線上生效。
// 貼上並部署後，請在編輯器手動執行一次 syncToCalendar，把已經殘留的舊日曆事件清乾淨。
// ============================================================

const SHEET_NAME = "訂單清單";
const CAL_NAME   = "🌸 花店訂單";
const ORDER_ID_COL = 13; // M 欄

function getConfig_() {
  const props = PropertiesService.getScriptProperties();
  return {
    apiSecret:     props.getProperty('API_SECRET'),
    adminPassword: props.getProperty('ADMIN_PASSWORD'),
  };
}

const STALLS = {
  "2026-05-23":["台科大"],
  "2026-05-30":["台大","師大","政大","北教大"],
  "2026-06-01":["協和佑德"],
  "2026-06-02":["南湖高中"],
  "2026-06-03":["南港國中"],
  "2026-06-06":["淡江","輔仁","海洋","北市大"],
  "2026-06-09":["興雅國中"],
  "2026-06-13":["東吳","實踐"],
  "2026-06-14":["東吳","實踐"],
  "2026-06-15":["福星國小"]
};

function getOrCreateCalendar() {
  const cals = CalendarApp.getCalendarsByName(CAL_NAME);
  if (cals.length > 0) return cals[0];
  return CalendarApp.createCalendar(CAL_NAME, {
    color: CalendarApp.Color.PINK,
    summary: "花光光花店訂單與出攤行程"
  });
}

function doGet(e)  { return handleRequest(e); }
function doPost(e) { return handleRequest(e); }

function handleRequest(e) {
  const output = ContentService.createTextOutput();
  output.setMimeType(ContentService.MimeType.JSON);
  try {
    const params = e.parameter || {};
    const config = getConfig_();
    const action = params.action;

    if (action === "login") {
      if (params.password === config.adminPassword) {
        output.setContent(JSON.stringify({ success: true, token: config.apiSecret }));
      } else {
        output.setContent(JSON.stringify({ success: false, message: "密碼錯誤" }));
      }
      return output;
    }

    if (params.secret !== config.apiSecret) {
      output.setContent(JSON.stringify({ success: false, message: "未授權" }));
      return output;
    }
    let result;
    if      (action === "getOrders")   result = getOrders();
    else if (action === "addOrder")    result = addOrder(JSON.parse(params.data || "{}"));
    else if (action === "updateOrder") result = updateOrder(Number(params.row), JSON.parse(params.data || "{}"));
    else if (action === "deleteOrder") result = deleteOrder(Number(params.row));
    else if (action === "markPaid")    result = markPaid(Number(params.row), params.value);
    else result = { success: false, message: "未知的 action" };
    output.setContent(JSON.stringify(result));
  } catch (err) {
    output.setContent(JSON.stringify({ success: false, message: err.message }));
  }
  return output;
}

function getOrders() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
  ensureNoteHeader_(sheet);
  const data  = sheet.getDataRange().getValues();
  const orders = [];
  for (let i = 1; i < data.length; i++) {
    const row  = data[i];
    const name = String(row[0] || "").trim();
    const date = formatDateCell(row[3]);
    if (!name && !date) continue;
    if (name.includes("成本") || name.includes("完成線")) continue;
    if (String(row[2]).trim().startsWith("-")) continue;
    orders.push({
      row:         i + 1,
      name:        name,
      phone:       restorePhone_(row[1]),
      amount:      row[2] || "",
      pickup_date: date,
      pickup_time: formatTimeCell(row[4]),
      delivery:    String(row[5] || ""),
      item:        String(row[6] || ""),
      last5:       String(row[7] || ""),
      paid:        row[8] === true || String(row[8]).toLowerCase() === "true" || row[8] === "✓",
      pickup_info: String(row[9] || ""),
      card_text:   String(row[10] || ""),
      // 舊訂單沒有 L 欄時就是空字串，不影響原本的列
      note:        String(row[11] || ""),
    });
  }
  return { success: true, orders };
}

// ── Date 物件轉 年/月/日。只丟月/日的話，前端會把每一年都當成今年。──
function formatDateCell(val) {
  if (!val) return "";
  if (val instanceof Date) {
    const y = val.getFullYear();
    const m = val.getMonth() + 1;
    const d = val.getDate();
    return `${y}/${m}/${d}`;
  }
  return String(val).trim();
}

// ── 時間欄位格式化 ──
function formatTimeCell(val) {
  if (!val) return "";
  if (val instanceof Date) {
    const h = val.getHours();
    const m = val.getMinutes();
    return `${h}:${String(m).padStart(2,'0')}`;
  }
  return String(val).trim();
}

function addOrder(data) {
  const sheet   = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
  const allData = sheet.getDataRange().getValues();
  const newDate = parseDate(data.pickup_date || "");

  let insertRow = allData.length + 1;
  for (let i = 1; i < allData.length; i++) {
    const existingDate = parseDate(formatDateCell(allData[i][3]));
    if (newDate && existingDate && newDate < existingDate) { insertRow = i + 1; break; }
  }

  sheet.insertRowBefore(insertRow);
  // 插入列時如果帶進別列的訂單編號，清掉，避免兩張單共用同一個日曆事件
  sheet.getRange(insertRow, ORDER_ID_COL).clearContent();
  writeOrderRow_(sheet, insertRow, [
    data.name        || "",
    data.phone       || "",
    data.amount      || "",
    data.pickup_date || "",
    data.pickup_time || "",
    data.delivery    || "",
    data.item        || "",
    data.last5       || "",
    false,
    data.pickup_info || "",
    data.card_text   || "",
    data.note        || "",
  ]);
  // 列已經寫進試算表後再對日曆，這樣新訂單才會被讀到
  SpreadsheetApp.flush();
  syncCalendarQuiet_();
  return { success: true, message: "訂單已新增於第 " + insertRow + " 行", row: insertRow };
}

function updateOrder(rowNumber, data) {
  if (!rowNumber || rowNumber < 2) return { success: false, message: "無效的行號" };
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
  writeOrderRow_(sheet, rowNumber, [
    data.name        || "",
    data.phone       || "",
    data.amount      || "",
    data.pickup_date || "",
    data.pickup_time || "",
    data.delivery    || "",
    data.item        || "",
    data.last5       || "",
    sheet.getRange(rowNumber, 9).getValue(),
    data.pickup_info || "",
    data.card_text   || "",
    data.note        || "",
  ]);
  SpreadsheetApp.flush();
  syncCalendarQuiet_();
  return { success: true };
}

// L1 是空白才寫上「備註」。已經有標題就不覆蓋。
function ensureNoteHeader_(sheet) {
  const cell = sheet.getRange(1, 12);
  if (!String(cell.getValue() || "").trim()) cell.setValue("備註");
}

// 電話、LINE Pay 交易代碼先設成純文字再寫入，試算表才不會把 0972... 存成數字
function writeOrderRow_(sheet, rowNumber, values) {
  const phone = restorePhone_(values[1]);
  const code  = String(values[7] == null ? "" : values[7]).trim();
  const phoneCell = sheet.getRange(rowNumber, 2);
  const codeCell  = sheet.getRange(rowNumber, 8);
  phoneCell.setNumberFormat("@");
  codeCell.setNumberFormat("@");
  const row = values.slice();
  row[1] = phone;
  row[7] = code;
  ensureNoteHeader_(sheet);
  sheet.getRange(rowNumber, 1, 1, 12).setValues([row]);
  // setValues 仍可能把長得像數字的字串轉成數字，文字格式設好後再寫一次這兩格
  phoneCell.setValue(phone);
  codeCell.setValue(code);
}

// 舊資料若已變成 9 開頭的 9 位數，讀出來時補回開頭的 0。
// 試算表把電話存成數字時，0972915271 會變成 972915271，字串化有時還會變成 9.72915271E8。
function restorePhone_(val) {
  if (val instanceof Date) return "";
  let s = "";
  if (typeof val === "number" && isFinite(val)) {
    s = Math.round(val).toFixed(0);
  } else {
    s = String(val == null ? "" : val).trim();
    if (/^9(?:\.\d+)?[eE]\+?0*8$/.test(s)) {
      const n = Number(s);
      if (isFinite(n)) s = Math.round(n).toFixed(0);
    }
  }
  if (/^\d+\.0+$/.test(s)) s = s.replace(/\.0+$/, "");
  if (/^9\d{8}$/.test(s)) return "0" + s;
  return s;
}

function deleteOrder(rowNumber) {
  if (!rowNumber || rowNumber < 2) return { success: false, message: "無效的行號" };
  SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME).deleteRow(rowNumber);
  SpreadsheetApp.flush();
  syncCalendarQuiet_();
  return { success: true };
}

function markPaid(rowNumber, value) {
  if (!rowNumber || rowNumber < 2) return { success: false, message: "無效的行號" };
  SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME)
    .getRange(rowNumber, 9).setValue(value === "true");
  return { success: true };
}

// 日曆同步失敗時，訂單還是要算儲存成功，下次再對一次就好
function syncCalendarQuiet_() {
  try {
    syncToCalendar();
  } catch (err) {
    Logger.log('日曆同步失敗（訂單資料已儲存）：' + (err && err.message ? err.message : err));
  }
}

function makeOrderId_() {
  // 開頭加 o，試算表才不會把編號當成數字，16 位數字會被吃掉
  return 'o' + Utilities.getUuid().replace(/-/g, '').slice(0, 16);
}

function cleanOrderId_(raw) {
  const s = String(raw == null ? '' : raw).trim();
  if (/^o[a-f0-9]{16}$/.test(s)) return s;
  return '';
}

function dateFromKey_(day) {
  const p = String(day || '').split('-');
  return new Date(parseInt(p[0], 10), parseInt(p[1], 10) - 1, parseInt(p[2], 10));
}

function ensureOrderIdHeader_(sheet) {
  const cell = sheet.getRange(1, ORDER_ID_COL);
  if (!String(cell.getValue() || '').trim()) cell.setValue('訂單編號');
}

function writeOrderIds_(sheet, lastRow, idByRow) {
  const keys = Object.keys(idByRow);
  if (!keys.length || lastRow < 2) return;
  const range = sheet.getRange(2, ORDER_ID_COL, lastRow - 1, 1);
  const values = range.getValues();
  for (let i = 0; i < values.length; i++) {
    const rowNumber = i + 2;
    if (idByRow[rowNumber]) values[i][0] = idByRow[rowNumber];
  }
  range.setNumberFormat('@');
  range.setValues(values);
}

function plainEvent_(ev) {
  const allDay = ev.isAllDayEvent();
  const start = allDay ? ev.getAllDayStartDate() : ev.getStartTime();
  const end = allDay ? null : ev.getEndTime();
  return {
    id: ev.getId(),
    title: ev.getTitle() || '',
    description: ev.getDescription() || '',
    day: dayKeyFromDate_(start),
    allDay: allDay,
    startH: end ? start.getHours() : null,
    startM: end ? start.getMinutes() : null,
    endH: end ? end.getHours() : null,
    endM: end ? end.getMinutes() : null,
  };
}

function writeOrderOntoEvent_(event, order) {
  const content = orderEventContent_(order);
  if ((event.getTitle() || '') !== content.title) event.setTitle(content.title);
  if (normText_(event.getDescription()) !== normText_(content.description)) {
    event.setDescription(content.description);
  }
  const date = dateFromKey_(content.day);
  if (content.allDay) {
    // 原本有時間、現在沒填時間時，會改成全日
    event.setAllDayDate(date);
  } else {
    const start = new Date(date.getFullYear(), date.getMonth(), date.getDate(), content.startH, content.startM, 0);
    const end = new Date(date.getFullYear(), date.getMonth(), date.getDate(), content.endH, content.endM, 0);
    event.setTime(start, end);
  }
}

function createOrderEvent_(cal, order) {
  const content = orderEventContent_(order);
  const date = dateFromKey_(content.day);
  let event;
  if (content.allDay) {
    event = cal.createAllDayEvent(content.title, date, { description: content.description });
  } else {
    const start = new Date(date.getFullYear(), date.getMonth(), date.getDate(), content.startH, content.startM, 0);
    const end = new Date(date.getFullYear(), date.getMonth(), date.getDate(), content.endH, content.endM, 0);
    event = cal.createEvent(content.title, start, end, { description: content.description });
  }
  event.addPopupReminder(60 * 24);
  event.addPopupReminder(60);
  return event;
}

function tryCalendar_(label, fn) {
  try {
    fn();
    return true;
  } catch (err) {
    Logger.log(label + '：' + (err && err.message ? err.message : err));
    return false;
  }
}

// 以前要手動執行 syncToCalendar（或自己設的時間觸發）才會同步。
// 現在新增、修改、刪除訂單也會呼叫它，所以一改取花日期，日曆就會跟著改。
// 每張訂單只留一筆最新事件：日期、時間、品項、金額變了就改同一筆；訂單刪了就把事件刪掉。
// 上個月 1 號之後的訂單會建立或更新行程。再往前大約一年半，只清掉已經沒有對應訂單的殘留事件，
// 試算表上還在的舊訂單不會被刪、也不會補建。🎪 出攤、以及不是這個系統建立的事件，不會動。
function syncToCalendar() {
  const cal   = getOrCreateCalendar();
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
  const data  = sheet.getDataRange().getValues();
  const now   = new Date();
  const windowStart = new Date(now.getFullYear(), now.getMonth() - 1, 1);
  const scanStart = new Date(windowStart.getFullYear(), windowStart.getMonth() - 18, 1);

  ensureOrderIdHeader_(sheet);

  const orders = [];
  const keepOrders = [];
  const extraLiveIds = [];
  const idByRow = {};
  const seenIds = {};

  for (let i = 1; i < data.length; i++) {
    const row        = data[i];
    const name       = String(row[0] || '').trim();
    const pickupDate = formatDateCell(row[3]);
    if (!name || !pickupDate) continue;
    if (name.includes('成本') || name.includes('完成線')) continue;

    const eventDate = parseDate(pickupDate);
    const existingId = cleanOrderId_(row[ORDER_ID_COL - 1]);
    if (!eventDate || eventDate < scanStart) {
      if (existingId) extraLiveIds.push(existingId);
      continue;
    }

    // 上個月之前的訂單：行程留著，但不要因為還沒建過就補一大堆舊行程
    if (eventDate < windowStart) {
      if (existingId) extraLiveIds.push(existingId);
      keepOrders.push({
        title: orderEventTitle_(name, String(row[6] || '').trim()),
        day: dayKeyFromDate_(eventDate),
      });
      continue;
    }

    // 同一列永遠用同一個編號。列被刪掉後，別列的編號會跟著那一列走，不會錯位。
    let orderId = existingId;
    if (!orderId || seenIds[orderId]) {
      orderId = makeOrderId_();
      idByRow[i + 1] = orderId;
    }
    seenIds[orderId] = true;

    orders.push({
      orderId:    orderId,
      name:       name,
      phone:      restorePhone_(row[1]),
      amount:     row[2] || '',
      delivery:   String(row[5] || '').trim(),
      item:       String(row[6] || '').trim(),
      pickupInfo: String(row[9] || '').trim(),
      pickupTime: formatTimeCell(row[4]),
      day:        dayKeyFromDate_(eventDate),
    });
  }

  writeOrderIds_(sheet, data.length, idByRow);

  let windowEnd = new Date(now.getFullYear() + 2, now.getMonth(), now.getDate() + 1);
  for (let i = 0; i < orders.length; i++) {
    const d = dateFromKey_(orders[i].day);
    const limit = new Date(d.getFullYear(), d.getMonth(), d.getDate() + 1);
    if (limit > windowEnd) windowEnd = limit;
  }

  const rawEvents = cal.getEvents(scanStart, windowEnd);
  const plain = [];
  const byId = {};
  for (let i = 0; i < rawEvents.length; i++) {
    const item = plainEvent_(rawEvents[i]);
    if (!item.id || byId[item.id]) continue;
    plain.push(item);
    byId[item.id] = rawEvents[i];
  }

  const plan = planOrderCalendarSync_(orders, plain, keepOrders, extraLiveIds);
  let created = 0, updated = 0, removed = 0;

  for (let i = 0; i < plan.remove.length; i++) {
    const ev = byId[plan.remove[i]];
    if (!ev) continue;
    if (tryCalendar_('刪除日曆事件失敗', function () { ev.deleteEvent(); })) removed++;
  }
  for (let i = 0; i < plan.update.length; i++) {
    const ev = byId[plan.update[i].eventId];
    if (!ev) continue;
    const order = plan.update[i].order;
    if (tryCalendar_('更新日曆事件失敗', function () { writeOrderOntoEvent_(ev, order); })) updated++;
  }
  for (let i = 0; i < plan.create.length; i++) {
    const order = plan.create[i];
    if (tryCalendar_('新增日曆事件失敗', function () { createOrderEvent_(cal, order); })) created++;
  }

  for (const [dateStr, venues] of Object.entries(STALLS)) {
    const parts = dateStr.split('-');
    const stallDate = new Date(parseInt(parts[0]), parseInt(parts[1]) - 1, parseInt(parts[2]));
    if (stallDate < windowStart) continue;
    const stallTitle = `🎪 出攤｜${venues.join('、')}`;
    const nextDay = new Date(stallDate); nextDay.setDate(nextDay.getDate() + 1);
    const existing = cal.getEvents(stallDate, nextDay).filter(ev => ev.getTitle() === stallTitle);
    if (existing.length > 0) continue;
    const ev = cal.createAllDayEvent(stallTitle, stallDate, { description: `出攤地點：${venues.join('、')}` });
    ev.addPopupReminder(60 * 24);
    created++;
  }

  Logger.log(`日曆同步：新增 ${created}、更新 ${updated}、刪除 ${removed}、已是最新 ${plan.skip}`);
  return { created, updated, removed, skipped: plan.skip };
}

/*__CALENDAR_SYNC_START__*/
function normText_(s) {
  return String(s || '').replace(/\r\n/g, '\n').replace(/\r/g, '\n').replace(/[ \t]+$/gm, '').replace(/\n+$/g, '');
}

function dayKeyFromDate_(d) {
  const y = d.getFullYear();
  const m = d.getMonth() + 1;
  const day = d.getDate();
  const mm = m < 10 ? '0' + m : String(m);
  const dd = day < 10 ? '0' + day : String(day);
  return y + '-' + mm + '-' + dd;
}

function orderIdFromDescription_(text) {
  const lines = normText_(text).split('\n');
  for (let i = lines.length - 1; i >= 0; i--) {
    const found = lines[i].match(/^【花店訂單：([^】]+)】$/);
    if (found) return found[1].trim();
  }
  return '';
}

function isStallEventTitle_(title) {
  return String(title || '').indexOf('🎪 出攤') === 0;
}

// 舊版同步沒有訂單編號。標題是「🌸 姓名｜品項」，說明從「客戶：」開始，才算這個系統建的。
function isLegacyOrderEvent_(title, description) {
  const t = String(title || '');
  if (t.indexOf('🌸 ') !== 0) return false;
  if (t.indexOf('｜') < 0) return false;
  if (orderIdFromDescription_(description)) return false;
  return normText_(description).indexOf('客戶：') === 0;
}

function orderEventTitle_(name, item) {
  const shortItem = String(item || '').replace(/《\s*/g, '').replace(/\s*》/g, '').trim().slice(0, 15);
  return '🌸 ' + String(name || '').trim() + '｜' + shortItem;
}

function orderEventDescription_(order) {
  const lines = ['客戶：' + (order.name || '')];
  if (order.phone) lines.push('電話：' + order.phone);
  if (order.amount) lines.push('金額：$' + order.amount);
  if (order.delivery) lines.push('配送：' + order.delivery);
  if (order.item) lines.push('品項：' + order.item);
  if (order.pickupInfo) lines.push('收花資訊：' + order.pickupInfo);
  lines.push('【花店訂單：' + order.orderId + '】');
  return lines.join('\n');
}

function resolveTime(timeStr) {
  if (!timeStr) return null;
  const s = timeStr.trim();
  const exact = s.match(/(\d{1,2})[:.：](\d{2})/);
  if (exact) {
    const h = parseInt(exact[1]), m = parseInt(exact[2]);
    return { startH: h, startM: m, endH: h, endM: m+30>59?59:m+30 };
  }
  if (/早上|上午/.test(s))           return { startH:9,  startM:0, endH:11, endM:0 };
  if (/接近中午|中午|午/.test(s))     return { startH:11, startM:0, endH:13, endM:0 };
  if (/下午/.test(s))                 return { startH:13, startM:0, endH:17, endM:0 };
  if (/傍晚/.test(s))                 return { startH:17, startM:0, endH:19, endM:0 };
  if (/晚上|晚/.test(s))              return { startH:19, startM:0, endH:21, endM:0 };
  return null;
}

function orderEventContent_(order) {
  const time = resolveTime(order.pickupTime || '');
  return {
    title: orderEventTitle_(order.name, order.item),
    description: orderEventDescription_(order),
    day: order.day,
    allDay: !time,
    startH: time ? time.startH : null,
    startM: time ? time.startM : null,
    endH: time ? time.endH : null,
    endM: time ? time.endM : null,
  };
}

function eventContentMatches_(ev, content) {
  if (ev.title !== content.title) return false;
  if (normText_(ev.description) !== normText_(content.description)) return false;
  if (ev.day !== content.day) return false;
  if (!!ev.allDay !== !!content.allDay) return false;
  if (!content.allDay) {
    if (Number(ev.startH) !== Number(content.startH)) return false;
    if (Number(ev.startM) !== Number(content.startM)) return false;
    if (Number(ev.endH) !== Number(content.endH)) return false;
    if (Number(ev.endM) !== Number(content.endM)) return false;
  }
  return true;
}

function pickBestEvent_(events, content) {
  let best = events[0];
  for (let i = 0; i < events.length; i++) {
    const ev = events[i];
    if (eventContentMatches_(ev, content)) return ev;
    if (ev.day === content.day && best.day !== content.day) best = ev;
  }
  return best;
}

function legacyBelongsToKept_(ev, kept) {
  for (let i = 0; i < kept.length; i++) {
    if (ev.day === kept[i].day && ev.title === kept[i].title) return true;
  }
  return false;
}

// 算出要新增、更新、刪除哪些事件。同一筆訂單只會對到一個事件。
// keepOrders：試算表還在、但太舊所以不重建的訂單。extraLiveIds：這些編號的事件不要當孤兒刪掉。
function planOrderCalendarSync_(orders, events, keepOrders, extraLiveIds) {
  const linkedById = {};
  const legacy = [];
  for (let i = 0; i < events.length; i++) {
    const ev = events[i];
    const oid = orderIdFromDescription_(ev.description);
    if (oid) {
      if (!linkedById[oid]) linkedById[oid] = [];
      linkedById[oid].push(ev);
      continue;
    }
    if (isStallEventTitle_(ev.title)) continue;
    if (isLegacyOrderEvent_(ev.title, ev.description)) legacy.push(ev);
  }

  const claimed = {};
  const kept = keepOrders || [];
  for (let i = 0; i < legacy.length; i++) {
    if (legacyBelongsToKept_(legacy[i], kept)) claimed[legacy[i].id] = true;
  }

  const create = [];
  const update = [];
  const remove = [];
  let skip = 0;
  const needLegacy = [];

  for (let i = 0; i < orders.length; i++) {
    const order = orders[i];
    const content = orderEventContent_(order);
    const ownedAll = linkedById[order.orderId] || [];
    const owned = [];
    for (let j = 0; j < ownedAll.length; j++) {
      if (!claimed[ownedAll[j].id]) owned.push(ownedAll[j]);
    }
    if (owned.length) {
      const keep = pickBestEvent_(owned, content);
      for (let j = 0; j < ownedAll.length; j++) {
        const ev = ownedAll[j];
        if (claimed[ev.id]) continue;
        claimed[ev.id] = true;
        if (ev.id !== keep.id) remove.push(ev.id);
      }
      if (eventContentMatches_(keep, content)) skip++;
      else update.push({ eventId: keep.id, order: order });
      continue;
    }
    needLegacy.push({ order: order, content: content });
  }

  // 舊事件沒有編號：先認「同一天、同一個標題」的那筆
  const still = [];
  for (let i = 0; i < needLegacy.length; i++) {
    const item = needLegacy[i];
    const hits = [];
    for (let j = 0; j < legacy.length; j++) {
      const ev = legacy[j];
      if (claimed[ev.id]) continue;
      if (ev.title === item.content.title && ev.day === item.content.day) hits.push(ev);
    }
    if (!hits.length) { still.push(item); continue; }
    const keep = pickBestEvent_(hits, item.content);
    claimed[keep.id] = true;
    if (eventContentMatches_(keep, item.content)) skip++;
    else update.push({ eventId: keep.id, order: item.order });
  }

  // 日期被改掉、新的那天還沒建事件時，通常只會剩下一筆同標題的舊事件，把它搬走
  const still2 = [];
  for (let i = 0; i < still.length; i++) {
    const item = still[i];
    const hits = [];
    for (let j = 0; j < legacy.length; j++) {
      const ev = legacy[j];
      if (claimed[ev.id]) continue;
      if (ev.title === item.content.title) hits.push(ev);
    }
    if (hits.length !== 1) { still2.push(item); continue; }
    claimed[hits[0].id] = true;
    if (eventContentMatches_(hits[0], item.content)) skip++;
    else update.push({ eventId: hits[0].id, order: item.order });
  }

  for (let i = 0; i < still2.length; i++) create.push(still2[i].order);

  const live = {};
  for (let i = 0; i < orders.length; i++) live[orders[i].orderId] = true;
  const extras = extraLiveIds || [];
  for (let i = 0; i < extras.length; i++) live[extras[i]] = true;
  const ids = Object.keys(linkedById);
  for (let i = 0; i < ids.length; i++) {
    if (live[ids[i]]) continue;
    const list = linkedById[ids[i]];
    for (let j = 0; j < list.length; j++) {
      if (!claimed[list[j].id]) remove.push(list[j].id);
    }
  }

  // 沒被任何訂單認領的舊訂單事件，就是改期或刪單留下的，刪掉
  for (let i = 0; i < legacy.length; i++) {
    if (!claimed[legacy[i].id]) remove.push(legacy[i].id);
  }

  return { create: create, update: update, remove: remove, skip: skip };
}
/*__CALENDAR_SYNC_END__*/

// 有西元年就用該年；只寫月/日才用今年。不要再寫死 2026，否則「2026/10/17」會被看成 26 月。
function parseDate(dateStr) {
  if (!dateStr) return null;
  const s = String(dateStr).trim();
  let m = s.match(/^(\d{4})[\/\-](\d{1,2})[\/\-](\d{1,2})/);
  if (m) return new Date(parseInt(m[1], 10), parseInt(m[2], 10) - 1, parseInt(m[3], 10));
  m = s.match(/^(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{4})$/);
  if (m) return new Date(parseInt(m[3], 10), parseInt(m[1], 10) - 1, parseInt(m[2], 10));
  m = s.match(/^(\d{1,2})[\/\-](\d{1,2})$/);
  if (m) return new Date(new Date().getFullYear(), parseInt(m[1], 10) - 1, parseInt(m[2], 10));
  return null;
}

function testSync()      { Logger.log(JSON.stringify(syncToCalendar())); }
function testGetOrders() { Logger.log(JSON.stringify(getOrders())); }