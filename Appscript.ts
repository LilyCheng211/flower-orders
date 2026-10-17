// ============================================================
// 花店訂單管理系統 - Apps Script 完整版 v2
// 欄位：A姓名/B電話/C金額/D取花日期/E幾點取花/F配送/G品項/H LINE Pay 交易代碼/I收到款/J收花資訊/K卡片內容
// H 欄在前端 JSON 仍叫 last5。電話與交易代碼要以文字寫入，避免開頭的 0 被吃掉。
// 這份檔案要貼回 Apps Script，並建立新的部署版本，才會在線上生效。
// ============================================================

const SHEET_NAME = "訂單清單";
const CAL_NAME   = "🌸 花店訂單";

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
  ]);
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
  ]);
  return { success: true };
}

// 電話、LINE Pay 交易代碼先設成純文字再寫入，試算表才不會把 0972... 存成數字
function writeOrderRow_(sheet, rowNumber, values) {
  const phone = String(values[1] == null ? "" : values[1]).trim();
  const code  = String(values[7] == null ? "" : values[7]).trim();
  const phoneCell = sheet.getRange(rowNumber, 2);
  const codeCell  = sheet.getRange(rowNumber, 8);
  phoneCell.setNumberFormat("@");
  codeCell.setNumberFormat("@");
  const row = values.slice();
  row[1] = phone;
  row[7] = code;
  sheet.getRange(rowNumber, 1, 1, 11).setValues([row]);
  // setValues 仍可能把長得像數字的字串轉成數字，文字格式設好後再寫一次這兩格
  phoneCell.setValue(phone);
  codeCell.setValue(code);
}

// 舊資料若已變成 9 開頭的 9 位數，讀出來時補回開頭的 0
function restorePhone_(val) {
  if (val instanceof Date) return "";
  let s = String(val == null ? "" : val).trim();
  if (/^\d+\.0+$/.test(s)) s = s.replace(/\.0+$/, "");
  if (/^9\d{8}$/.test(s)) return "0" + s;
  return s;
}

function deleteOrder(rowNumber) {
  if (!rowNumber || rowNumber < 2) return { success: false, message: "無效的行號" };
  SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME).deleteRow(rowNumber);
  return { success: true };
}

function markPaid(rowNumber, value) {
  if (!rowNumber || rowNumber < 2) return { success: false, message: "無效的行號" };
  SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME)
    .getRange(rowNumber, 9).setValue(value === "true");
  return { success: true };
}

function syncToCalendar() {
  const cal   = getOrCreateCalendar();
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
  const data  = sheet.getDataRange().getValues();
  const now   = new Date();
  let created = 0, skipped = 0;

  for (let i = 1; i < data.length; i++) {
    const row        = data[i];
    const name       = String(row[0] || "").trim();
    const pickupDate = formatDateCell(row[3]);
    const pickupTime = formatTimeCell(row[4]);
    const item       = String(row[6] || "").trim();
    const amount     = row[2] || "";
    const delivery   = String(row[5] || "").trim();
    const phone      = restorePhone_(row[1]);
    const pickupInfo = String(row[9] || "").trim();

    if (!name || !pickupDate) continue;
    if (name.includes("成本") || name.includes("完成線")) continue;

    const eventDate = parseDate(pickupDate);
    if (!eventDate || eventDate < new Date(now.getFullYear(), now.getMonth()-1, 1)) continue;

    const shortItem = item.replace(/《\s*/g,'').replace(/\s*》/g,'').trim().slice(0,15);
    const title = `🌸 ${name}｜${shortItem}`;
    const nextDay = new Date(eventDate); nextDay.setDate(nextDay.getDate()+1);
    const existing = cal.getEvents(eventDate, nextDay).filter(ev => ev.getTitle() === title);
    if (existing.length > 0) { skipped++; continue; }

    const description = [
      `客戶：${name}`,
      phone      ? `電話：${phone}` : '',
      amount     ? `金額：$${amount}` : '',
      delivery   ? `配送：${delivery}` : '',
      item       ? `品項：${item}` : '',
      pickupInfo ? `收花資訊：${pickupInfo}` : '',
    ].filter(Boolean).join('\n');

    const timeRange = resolveTime(pickupTime);
    let event;
    if (timeRange) {
      const start = new Date(eventDate);
      start.setHours(timeRange.startH, timeRange.startM, 0);
      const end = new Date(eventDate);
      end.setHours(timeRange.endH, timeRange.endM, 0);
      event = cal.createEvent(title, start, end, { description });
    } else {
      event = cal.createAllDayEvent(title, eventDate, { description });
    }
    event.addPopupReminder(60 * 24);
    event.addPopupReminder(60);
    created++;
  }

  for (const [dateStr, venues] of Object.entries(STALLS)) {
    const parts = dateStr.split('-');
    const stallDate = new Date(parseInt(parts[0]), parseInt(parts[1])-1, parseInt(parts[2]));
    if (stallDate < new Date(now.getFullYear(), now.getMonth()-1, 1)) continue;
    const stallTitle = `🎪 出攤｜${venues.join('、')}`;
    const nextDay = new Date(stallDate); nextDay.setDate(nextDay.getDate()+1);
    const existing = cal.getEvents(stallDate, nextDay).filter(ev => ev.getTitle() === stallTitle);
    if (existing.length > 0) continue;
    const ev = cal.createAllDayEvent(stallTitle, stallDate, { description: `出攤地點：${venues.join('、')}` });
    ev.addPopupReminder(60 * 24);
    created++;
  }

  Logger.log(`Calendar 同步：新增 ${created} 筆，略過 ${skipped} 筆`);
  return { created, skipped };
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