// ============================================================
// 花店訂單管理系統 - Apps Script 完整版 v2
// 欄位：A姓名/B電話/C金額/D取花日期/E幾點取花/F配送/G品項/H LINE Pay 交易代碼/I收到款/J收花資訊/K卡片內容/L備註
// H 欄在前端 JSON 仍叫 last5。電話與交易代碼要以文字寫入，避免開頭的 0 被吃掉。
// L 欄備註用純文字寫入，換行要留著。不要把備註寫進 M、N 或 O。
// M 欄「訂單編號」是日曆同步自己寫的。刪掉一列時列號會往上移，所以不能用列號認訂單。
// N 欄「請求編號」是瀏覽器這一次送出產生的。同一編號再進來，只回應成功，不再新增一列，也不再上傳一次圖片。
// O 欄「參考圖片」是第 15 欄。一格可以有多個連結，一行一個，在試算表裡可以點開。
// 圖片放在雲端硬碟資料夾「花店訂單參考圖」。資料夾本身不分享，所以整包照片不會被搜尋到。
// 每一張圖另外設成「知道連結的使用者可檢視」，訂單頁的縮圖才顯示得了。拿到單一連結的人看得到那一張。
// 新增、修改、刪除只先寫試算表就回應。日曆整份對帳改排程，避免一次請求卡太久又被重送。
// 登入成功只發簽章憑證（版本|到期|HMAC），不把 API_SECRET 交給瀏覽器。憑證版本存在 Script Properties。
// 這份檔案要貼回現有的 Apps Script，用「部署」→「管理部署」編輯 API_URL 那個部署，版本選「新版本」。
// 不要新建另一個部署，網址才不會變。前端要等這個新版本上線後再合併。
// 貼上並部署後，請在編輯器手動執行一次 syncToCalendar，把已經殘留的舊日曆事件清乾淨。
// 也請手動執行一次 authorizeTriggers，同意排程權限。之後每次改訂單才會自動排日曆同步。
// 這版會用到雲端硬碟。請再手動執行一次 authorizeDrive，同意 Drive 權限：
// https://www.googleapis.com/auth/drive
// ============================================================

const SHEET_NAME = "訂單清單";
const CAL_NAME   = "🌸 花店訂單";
const ORDER_ID_COL = 13; // M 欄
const REQUEST_ID_COL = 14; // N 欄
const NOTE_COL = 12; // L 欄
const IMAGE_COL = 15; // O 欄，參考圖片連結
const IMAGE_FOLDER_NAME = "花店訂單參考圖";

function getConfig_() {
  const props = PropertiesService.getScriptProperties();
  return {
    apiSecret:     props.getProperty('API_SECRET'),
    adminPassword: props.getProperty('ADMIN_PASSWORD'),
  };
}

// 憑證版本沒寫過時當作 1。登出所有裝置才把這個數字加一，舊憑證就全部失效。
function currentTokenVersion_() {
  const raw = PropertiesService.getScriptProperties().getProperty("TOKEN_VERSION");
  const n = parseInt(raw, 10);
  if (!isFinite(n) || n < 1) return 1;
  return n;
}

// HMAC-SHA256 的位元組轉成 64 個小寫十六進位，前後端才對得起同一條簽章。
function hmacHex_(text, secret) {
  const bytes = Utilities.computeHmacSha256Signature(String(text), String(secret == null ? "" : secret));
  let hex = "";
  for (let i = 0; i < bytes.length; i++) {
    let n = bytes[i];
    if (n < 0) n += 256;
    const h = n.toString(16);
    hex += h.length < 2 ? "0" + h : h;
  }
  return hex;
}

function issueSessionToken_(config) {
  const version = currentTokenVersion_();
  const expiry = Date.now() + tokenTtlMs_();
  const payload = tokenPayload_(version, expiry);
  return payload + "|" + hmacHex_(payload, config.apiSecret);
}

function tokenIsValid_(token, config) {
  if (!config || !config.apiSecret) return false;
  const parsed = parseSessionToken_(token);
  if (!parsed) return false;
  const expected = hmacHex_(parsed.payload, config.apiSecret);
  return sessionTokenStatus_(token, Date.now(), currentTokenVersion_(), expected) === "ok";
}

function logoutAllDevices_() {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(10000)) return { success: false, message: "系統忙碌，請再試一次" };
  try {
    const next = currentTokenVersion_() + 1;
    PropertiesService.getScriptProperties().setProperty("TOKEN_VERSION", String(next));
    return { success: true };
  } finally {
    lock.releaseLock();
  }
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
    // 讀取仍走網址參數。寫入改走 POST 本文，備註的換行才不會在網址裡被截掉。
    const params = readParams_(e);
    const config = getConfig_();
    const action = params.action;

    if (action === "login") {
      if (params.password === config.adminPassword) {
        // 密碼比對維持原本的 adminPassword。通過後只發 30 天簽章憑證，不回傳 apiSecret。
        if (!config.apiSecret) {
          output.setContent(JSON.stringify({ success: false, message: "未授權" }));
        } else {
          output.setContent(JSON.stringify({ success: true, token: issueSessionToken_(config) }));
        }
      } else {
        output.setContent(JSON.stringify({ success: false, message: "密碼錯誤" }));
      }
      return output;
    }

    if (!tokenIsValid_(params.secret, config)) {
      output.setContent(JSON.stringify({ success: false, message: "未授權" }));
      return output;
    }
    let result;
    if      (action === "getOrders")   result = getOrders();
    else if (action === "addOrder")    result = addOrder(JSON.parse(params.data || "{}"), params.requestId);
    else if (action === "updateOrder") result = updateOrder(Number(params.row), JSON.parse(params.data || "{}"), params.requestId);
    else if (action === "deleteOrder") result = deleteOrder(Number(params.row), params.requestId, params.name);
    else if (action === "markPaid")    result = markPaid(Number(params.row), params.value);
    else if (action === "logoutAll")   result = logoutAllDevices_();
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
  ensureImageHeader_(sheet);
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return { success: true, orders: [] };
  // 固定讀到 O 欄。getDataRange 有時不會把剛加上的備註或參考圖片算進去。
  const width = Math.max(IMAGE_COL, sheet.getLastColumn());
  const data = sheet.getRange(1, 1, lastRow, width).getValues();
  const orders = [];
  for (let i = 1; i < data.length; i++) {
    const row  = data[i];
    const name = String(row[0] || "").trim();
    const date = formatDateCell(row[3]);
    if (!name && !date) continue;
    if (name.includes("成本") || name.includes("完成線")) continue;
    if (String(row[2]).trim().startsWith("-")) continue;
    const item = orderFromCells_(row, i + 1);
    item.phone = restorePhone_(row[1]);
    item.pickup_date = date;
    item.pickup_time = formatTimeCell(row[4]);
    orders.push(item);
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

function addOrder(data, requestId) {
  const id = cleanRequestId_(requestId || (data && data.requestId));
  const lock = LockService.getScriptLock();
  // 同一筆送出若被重送，要排好隊，不能兩個人同時插列
  if (!lock.tryLock(20000)) return { success: false, message: "系統忙碌，請再試一次" };
  let result;
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
    ensureRequestIdHeader_(sheet);
    if (id) {
      const remembered = rememberedRow_(sheet, id);
      if (remembered) {
        result = { success: true, duplicate: true, row: remembered, message: "這筆訂單已經新增過" };
      }
    }
    if (!result) result = insertOrder_(sheet, data || {}, id);
  } finally {
    lock.releaseLock();
  }
  if (result && result.success && !result.duplicate) scheduleCalendarSync_();
  return result;
}

function insertOrder_(sheet, data, requestId) {
  // 先處理圖片。上傳失敗就不要插列，同一個請求編號重送才不會變成兩列或兩套圖。
  const imageError = validateIncomingImages_(data && data.images);
  if (imageError) return { success: false, message: imageError };
  let imageText = "";
  try {
    imageText = saveOrderImages_(data || {}, requestId);
  } catch (err) {
    return { success: false, message: imageErrorMessage_(err) };
  }

  const lastRow = Math.max(sheet.getLastRow(), 1);
  const width = Math.max(4, sheet.getLastColumn());
  const allData = sheet.getRange(1, 1, lastRow, width).getValues();
  const newDate = parseDate(data.pickup_date || "");

  let insertRow = allData.length + 1;
  for (let i = 1; i < allData.length; i++) {
    const existingDate = parseDate(formatDateCell(allData[i][3]));
    if (newDate && existingDate && newDate < existingDate) { insertRow = i + 1; break; }
  }

  if (insertRow <= sheet.getLastRow()) sheet.insertRowBefore(insertRow);
  // 插入列時如果帶進別列的編號，清掉，避免兩張單共用同一個日曆事件或請求編號
  sheet.getRange(insertRow, ORDER_ID_COL).clearContent();
  sheet.getRange(insertRow, REQUEST_ID_COL).clearContent();
  // I 欄跟 markPaid 一樣寫布林。表單沒勾就是 false。
  const cells = orderToCells_(data, paidFlag_(data && data.paid));
  writeOrderRow_(sheet, insertRow, cells);
  // 編輯舊列時備註會留在 L。新插入的列有時寫不進最後一欄，所以新增這裡再寫一次備註。
  const createdNote = sheet.getRange(insertRow, NOTE_COL);
  createdNote.setNumberFormat("@");
  createdNote.setValue(cells[11]);
  writeImageCell_(sheet, insertRow, imageText);
  if (requestId) rememberRequestId_(sheet, insertRow, requestId);
  SpreadsheetApp.flush();
  return {
    success: true,
    message: "訂單已新增於第 " + insertRow + " 行",
    row: insertRow,
    images: parseImageCell_(imageText)
  };
}

function updateOrder(rowNumber, data, requestId) {
  if (!rowNumber || rowNumber < 2) return { success: false, message: "無效的行號" };
  const id = cleanRequestId_(requestId || (data && data.requestId));
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(20000)) return { success: false, message: "系統忙碌，請再試一次" };
  let result;
  try {
    if (id && CacheService.getScriptCache().get("req:" + id)) {
      result = { success: true, duplicate: true };
      // 上次已經寫入，只是瀏覽器逾時。把 O 欄現有的圖回傳，明細才不用再等手動刷新。
      try {
        const savedSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
        result.images = parseImageCell_(savedSheet.getRange(rowNumber, IMAGE_COL).getValue());
      } catch (err) {
        Logger.log("讀取已儲存的參考圖失敗：" + (err && err.message ? err.message : err));
      }
    } else {
      const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
      const paidNow = sheet.getRange(rowNumber, 9).getValue();
      // 舊版頁面不會傳 images。沒有這個欄位就不要清掉已經存好的參考圖。
      if (hasImageField_(data)) {
        const imageError = validateIncomingImages_(data.images);
        if (imageError) {
          result = { success: false, message: imageError };
        } else {
          let imageText = "";
          try {
            imageText = saveOrderImages_(data, id);
          } catch (err) {
            result = { success: false, message: imageErrorMessage_(err) };
          }
          if (!result) {
            const previousImages = String(sheet.getRange(rowNumber, IMAGE_COL).getValue() || "");
            writeOrderRow_(sheet, rowNumber, orderToCells_(data || {}, paidNow));
            writeImageCell_(sheet, rowNumber, imageText);
            try { trashRemovedImages_(previousImages, imageText); }
            catch (err) { Logger.log("移除參考圖失敗：" + (err && err.message ? err.message : err)); }
            if (id) CacheService.getScriptCache().put("req:" + id, "u" + rowNumber, 21600);
            SpreadsheetApp.flush();
            result = { success: true, images: parseImageCell_(imageText) };
          }
        }
      } else {
        writeOrderRow_(sheet, rowNumber, orderToCells_(data || {}, paidNow));
        if (id) CacheService.getScriptCache().put("req:" + id, "u" + rowNumber, 21600);
        SpreadsheetApp.flush();
        result = { success: true };
      }
    }
  } finally {
    lock.releaseLock();
  }
  if (result && result.success && !result.duplicate) scheduleCalendarSync_();
  return result;
}

/*__ORDER_IO_START__*/
function tokenTtlMs_() {
  return 30 * 24 * 60 * 60 * 1000;
}

function tokenPayload_(version, expiryMs) {
  return String(version) + "|" + String(expiryMs);
}

function parseSessionToken_(token) {
  const parts = String(token == null ? "" : token).split("|");
  if (parts.length !== 3) return null;
  if (!/^\d+$/.test(parts[0]) || !/^\d+$/.test(parts[1])) return null;
  if (!/^[a-f0-9]{64}$/.test(parts[2])) return null;
  return {
    version: Number(parts[0]),
    expiry: Number(parts[1]),
    signature: parts[2],
    payload: parts[0] + "|" + parts[1]
  };
}

// 簽章、到期、版本都要過。版本不同代表已經按過登出所有裝置。
function sessionTokenStatus_(token, nowMs, currentVersion, expectedSig) {
  const parsed = parseSessionToken_(token);
  if (!parsed) return "invalid";
  if (!(parsed.expiry > nowMs)) return "expired";
  if (parsed.version !== Number(currentVersion)) return "revoked";
  if (parsed.signature !== expectedSig) return "bad-signature";
  return "ok";
}

// 跟 markPaid 一樣：只有布林 true 或字串 "true"（以及舊的勾）算已收款。
function paidFlag_(value) {
  return value === true || String(value).toLowerCase() === "true" || value === "✓";
}

function normalizeNote_(val) {
  return String(val == null ? "" : val).replace(/\r\n/g, "\n").replace(/\r/g, "\n").trim();
}

function cleanRequestId_(raw) {
  const s = String(raw == null ? "" : raw).trim();
  if (/^r[a-z0-9]{6,40}$/.test(s)) return s;
  return "";
}

// ids[0] 是第 2 列。找到同一個請求編號就回傳列號，沒有就是 0。
function findRequestRow_(ids, requestId) {
  const id = cleanRequestId_(requestId);
  if (!id || !ids) return 0;
  for (let i = 0; i < ids.length; i++) {
    if (ids[i] === id) return i + 2;
  }
  return 0;
}

function decideDuplicate_(knownIds, requestId) {
  const row = findRequestRow_(knownIds, requestId);
  return { id: cleanRequestId_(requestId), duplicate: row > 0, row: row };
}

function orderToCells_(data, paid) {
  const src = data || {};
  return [
    src.name || "",
    src.phone || "",
    src.amount || "",
    src.pickup_date || "",
    src.pickup_time || "",
    src.delivery || "",
    src.item || "",
    src.last5 || "",
    paid,
    src.pickup_info || "",
    src.card_text || "",
    normalizeNote_(src.note),
  ];
}

function orderFromCells_(row, rowNumber) {
  const cells = row || [];
  return {
    row: rowNumber,
    name: String(cells[0] || "").trim(),
    phone: cells[1],
    amount: cells[2] || "",
    pickup_date: cells[3],
    pickup_time: cells[4],
    delivery: String(cells[5] || ""),
    item: String(cells[6] || ""),
    last5: String(cells[7] || ""),
    paid: cells[8] === true || String(cells[8]).toLowerCase() === "true" || cells[8] === "✓",
    pickup_info: String(cells[9] || ""),
    card_text: String(cells[10] || ""),
    note: normalizeNote_(cells[11]),
    // 第 15 欄（O）才是參考圖片。舊訂單沒有這一欄時就是空的。
    images: parseImageCell_(cells[14]),
  };
}

function maxOrderImages_() { return 6; }
// 一張圖的 base64 上限。前端會先壓到大約 450KB，這裡再擋一次，避免撐爆 Apps Script 請求。
function maxImageChars_() { return 700000; }

function hasImageField_(data) {
  return !!(data && Object.prototype.hasOwnProperty.call(data, "images"));
}

function imageFileId_(link) {
  const s = String(link || "");
  let found = s.match(/\/file\/d\/([a-zA-Z0-9_-]{10,})/);
  if (found) return found[1];
  found = s.match(/[?&]id=([a-zA-Z0-9_-]{10,})/);
  if (found) return found[1];
  found = s.match(/\/d\/([a-zA-Z0-9_-]{10,})/);
  if (found) return found[1];
  return "";
}

function imageOpenUrl_(fileId) {
  return "https://drive.google.com/file/d/" + fileId + "/view";
}

function cleanImageUrl_(raw) {
  const s = String(raw == null ? "" : raw).trim();
  if (!/^https:\/\/[^\s]+$/i.test(s)) return "";
  if (s.length > 500) return "";
  return s;
}

function canonicalImageUrl_(raw) {
  const url = cleanImageUrl_(raw);
  if (!url) return "";
  const id = imageFileId_(url);
  if (!id) return "";
  return imageOpenUrl_(id);
}

function parseImageCell_(raw) {
  if (raw == null) return [];
  const parts = String(raw).split(/\s+/);
  const out = [];
  const seen = {};
  for (let i = 0; i < parts.length; i++) {
    const url = canonicalImageUrl_(parts[i]);
    if (!url || seen[url]) continue;
    seen[url] = true;
    out.push(url);
  }
  return out;
}

function joinImageLinks_(urls) {
  const text = Array.isArray(urls) ? urls.join("\n") : urls;
  return parseImageCell_(text).join("\n");
}

function validateIncomingImages_(images) {
  if (images == null || images === "") return "";
  if (!Array.isArray(images)) return "參考圖片格式不正確";
  if (images.length > maxOrderImages_()) return "參考圖片最多 " + maxOrderImages_() + " 張";
  let total = 0;
  for (let i = 0; i < images.length; i++) {
    const item = images[i] || {};
    if (item.url || item.link) continue;
    const data = String(item.data || "").replace(/\s/g, "");
    if (!data) continue;
    if (!/^[A-Za-z0-9+/=]+$/.test(data)) return "參考圖片格式不正確";
    if (data.length > maxImageChars_()) return "有一張圖片太大，請減少張數或換一張";
    total += data.length;
  }
  if (total > maxImageChars_() * maxOrderImages_()) return "參考圖片太大，請減少張數";
  return "";
}

function normalizeIncomingImages_(images) {
  const list = Array.isArray(images) ? images : [];
  const items = [];
  const limit = maxOrderImages_();
  for (let i = 0; i < list.length && items.length < limit; i++) {
    const item = list[i] || {};
    const url = canonicalImageUrl_(item.url || item.link || "");
    if (url) {
      items.push({ kind: "keep", url: url });
      continue;
    }
    const data = String(item.data || "").replace(/\s/g, "");
    if (!data) continue;
    if (!/^[A-Za-z0-9+/=]+$/.test(data)) continue;
    if (data.length > maxImageChars_()) continue;
    items.push({ kind: "upload", data: data });
  }
  return items;
}

// 同一個請求編號、同一張序，檔名固定。重送時看到同名檔就不要再上傳一次。
function uploadFileName_(requestId, index) {
  const id = cleanRequestId_(requestId);
  if (!id) return "";
  return "ref-" + id + "-" + index + ".jpg";
}

function planImageUploads_(requestId, items, existingNames) {
  const have = {};
  const names = existingNames || [];
  for (let i = 0; i < names.length; i++) have[String(names[i])] = true;
  const plan = [];
  let uploadIndex = 0;
  const list = items || [];
  for (let i = 0; i < list.length; i++) {
    if (!list[i] || list[i].kind !== "upload") continue;
    const name = uploadFileName_(requestId, uploadIndex);
    plan.push({
      index: uploadIndex,
      name: name,
      reuse: !!(name && have[name])
    });
    uploadIndex++;
  }
  return plan;
}

function collectImageLinks_(items, fileIdByUploadIndex) {
  const links = [];
  let uploadIndex = 0;
  const ids = fileIdByUploadIndex || {};
  const list = items || [];
  for (let i = 0; i < list.length; i++) {
    const item = list[i];
    if (!item) continue;
    if (item.kind === "keep") links.push(item.url);
    else if (item.kind === "upload") {
      if (ids[uploadIndex]) links.push(imageOpenUrl_(ids[uploadIndex]));
      uploadIndex++;
    }
  }
  return joinImageLinks_(links);
}
/*__ORDER_IO_END__*/

// L1 是空白才寫上「備註」。已經有標題就不覆蓋。
function ensureNoteHeader_(sheet) {
  const cell = sheet.getRange(1, 12);
  if (!String(cell.getValue() || "").trim()) cell.setValue("備註");
}

// 電話、LINE Pay 交易代碼先設成純文字再寫入，試算表才不會把 0972... 存成數字
function writeOrderRow_(sheet, rowNumber, values) {
  const phone = restorePhone_(values[1]);
  const code  = String(values[7] == null ? "" : values[7]).trim();
  const note  = normalizeNote_(values[11]);
  const phoneCell = sheet.getRange(rowNumber, 2);
  const codeCell  = sheet.getRange(rowNumber, 8);
  const noteCell  = sheet.getRange(rowNumber, NOTE_COL);
  phoneCell.setNumberFormat("@");
  codeCell.setNumberFormat("@");
  // 先設成純文字，備註裡的換行、； 才不會被試算表改掉
  noteCell.setNumberFormat("@");
  const row = values.slice();
  row[1] = phone;
  row[7] = code;
  row[11] = note;
  ensureNoteHeader_(sheet);
  sheet.getRange(rowNumber, 1, 1, 12).setValues([row]);
  // setValues 仍可能把長得像數字的字串轉成數字，文字格式設好後再寫一次
  phoneCell.setValue(phone);
  codeCell.setValue(code);
  noteCell.setValue(note);
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

function deleteOrder(rowNumber, requestId, expectName) {
  if (!rowNumber || rowNumber < 2) return { success: false, message: "無效的行號" };
  const id = cleanRequestId_(requestId);
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(20000)) return { success: false, message: "系統忙碌，請再試一次" };
  let result;
  try {
    if (id && CacheService.getScriptCache().get("req:" + id)) {
      result = { success: true, duplicate: true };
    } else {
      const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
      const name = String(sheet.getRange(rowNumber, 1).getValue() || "").trim();
      const wanted = String(expectName == null ? "" : expectName).trim();
      // 列號會往上移。對不到原本的姓名就不要刪，避免刪到下一筆。
      if (wanted && name !== wanted) {
        result = { success: false, message: "這一行已經不是原本那筆訂單，請重新整理後再刪" };
      } else {
        const imageText = String(sheet.getRange(rowNumber, IMAGE_COL).getValue() || "");
        sheet.deleteRow(rowNumber);
        try { trashRemovedImages_(imageText, ""); }
        catch (err) { Logger.log("刪除參考圖失敗：" + (err && err.message ? err.message : err)); }
        if (id) CacheService.getScriptCache().put("req:" + id, "d" + rowNumber, 21600);
        SpreadsheetApp.flush();
        result = { success: true };
      }
    }
  } finally {
    lock.releaseLock();
  }
  if (result && result.success && !result.duplicate) scheduleCalendarSync_();
  return result;
}

function markPaid(rowNumber, value) {
  if (!rowNumber || rowNumber < 2) return { success: false, message: "無效的行號" };
  SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME)
    .getRange(rowNumber, 9).setValue(value === "true");
  return { success: true };
}

// 不在這次回應裡跑完整日曆對帳。對帳會掃很多行程，請求會變慢，慢的 GET 又容易被重送。
function scheduleCalendarSync_() {
  try {
    const handler = "runScheduledCalendarSync";
    const triggers = ScriptApp.getProjectTriggers();
    for (let i = 0; i < triggers.length; i++) {
      if (triggers[i].getHandlerFunction() === handler) return;
    }
    ScriptApp.newTrigger(handler).timeBased().after(60 * 1000).create();
  } catch (err) {
    Logger.log("排程日曆同步失敗（訂單資料已儲存）：" + (err && err.message ? err.message : err));
  }
}

function runScheduledCalendarSync() {
  const triggers = ScriptApp.getProjectTriggers();
  for (let i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === "runScheduledCalendarSync") {
      ScriptApp.deleteTrigger(triggers[i]);
    }
  }
  syncToCalendar();
}

// 部署後在編輯器執行一次，讓 Google 跳出排程授權。
function authorizeTriggers() {
  scheduleCalendarSync_();
  Logger.log("已排定日曆同步");
}

function readParams_(e) {
  const params = {};
  const query = (e && e.parameter) || {};
  const qKeys = Object.keys(query);
  for (let i = 0; i < qKeys.length; i++) params[qKeys[i]] = query[qKeys[i]];
  const raw = e && e.postData && e.postData.contents;
  if (!raw || raw.charAt(0) !== "{") return params;
  const body = JSON.parse(raw);
  const keys = Object.keys(body);
  for (let i = 0; i < keys.length; i++) {
    const key = keys[i];
    const value = body[key];
    if (value == null) continue;
    // data 在本文裡是物件，後面仍用 JSON.parse 讀，跟舊的網址參數同一條路
    if (key === "data" && typeof value === "object") params.data = JSON.stringify(value);
    else params[key] = String(value);
  }
  return params;
}

function rememberedRow_(sheet, requestId) {
  const cached = CacheService.getScriptCache().get("req:" + requestId);
  const cachedRow = Number(cached);
  if (cachedRow >= 2) return cachedRow;
  const row = findRequestRow_(listRequestIds_(sheet), requestId);
  if (row) CacheService.getScriptCache().put("req:" + requestId, String(row), 21600);
  return row;
}

function rememberRequestId_(sheet, rowNumber, requestId) {
  ensureRequestIdHeader_(sheet);
  const cell = sheet.getRange(rowNumber, REQUEST_ID_COL);
  cell.setNumberFormat("@");
  cell.setValue(requestId);
  CacheService.getScriptCache().put("req:" + requestId, String(rowNumber), 21600);
}

function listRequestIds_(sheet) {
  const last = sheet.getLastRow();
  if (last < 2) return [];
  const values = sheet.getRange(2, REQUEST_ID_COL, last - 1, 1).getValues();
  const ids = [];
  for (let i = 0; i < values.length; i++) ids.push(String(values[i][0] || ""));
  return ids;
}

function ensureRequestIdHeader_(sheet) {
  const cell = sheet.getRange(1, REQUEST_ID_COL);
  if (!String(cell.getValue() || "").trim()) cell.setValue("請求編號");
}

function ensureImageHeader_(sheet) {
  const cell = sheet.getRange(1, IMAGE_COL);
  if (!String(cell.getValue() || "").trim()) cell.setValue("參考圖片");
}

// O 欄存完整連結，一行一個，試算表上可以點。沒有圖就清掉，避免插入列時抄到上一列。
function writeImageCell_(sheet, rowNumber, text) {
  ensureImageHeader_(sheet);
  const cell = sheet.getRange(rowNumber, IMAGE_COL);
  const value = String(text || "");
  if (!value) {
    cell.clearContent();
    return;
  }
  cell.setNumberFormat("@");
  try {
    const urls = value.split("\n");
    const rich = SpreadsheetApp.newRichTextValue().setText(value);
    let offset = 0;
    for (let i = 0; i < urls.length; i++) {
      const url = urls[i];
      if (/^https:\/\//i.test(url)) rich.setLinkUrl(offset, offset + url.length, url);
      offset += url.length + 1;
    }
    cell.setRichTextValue(rich.build());
  } catch (err) {
    cell.setValue(value);
  }
}

function imageErrorMessage_(err) {
  const msg = String(err && err.message ? err.message : err || "");
  if (/permission|權限|Authorization|auth/i.test(msg)) {
    return "參考圖片需要雲端硬碟權限，請擁有者重新授權後再試";
  }
  return "參考圖片上傳失敗，請再試一次";
}

// 資料夾只放在擁有者的雲端硬碟，不對整包開放分享。
function getOrCreateImageFolder_() {
  const props = PropertiesService.getScriptProperties();
  const saved = props.getProperty("IMAGE_FOLDER_ID");
  if (saved) {
    try {
      const existing = DriveApp.getFolderById(saved);
      if (existing && !existing.isTrashed()) return existing;
    } catch (err) {
      Logger.log("找不到參考圖資料夾，重新建立：" + err);
    }
  }
  const folder = DriveApp.createFolder(IMAGE_FOLDER_NAME);
  props.setProperty("IMAGE_FOLDER_ID", folder.getId());
  return folder;
}

function findNamedImageFile_(folder, name) {
  if (!folder || !name) return null;
  const found = folder.getFilesByName(name);
  if (found.hasNext()) return found.next();
  return null;
}

// 回傳要寫進 O 欄的連結。沒有新圖時不會去碰雲端硬碟。
function saveOrderImages_(data, requestId) {
  const items = normalizeIncomingImages_(data && data.images);
  const uploads = [];
  for (let i = 0; i < items.length; i++) {
    if (items[i].kind === "upload") uploads.push(items[i]);
  }
  if (!uploads.length) {
    const kept = [];
    for (let i = 0; i < items.length; i++) if (items[i].url) kept.push(items[i].url);
    return joinImageLinks_(kept);
  }

  const folder = getOrCreateImageFolder_();
  const existing = [];
  for (let i = 0; i < uploads.length; i++) {
    const name = uploadFileName_(requestId, i);
    if (name && findNamedImageFile_(folder, name)) existing.push(name);
  }
  const plan = planImageUploads_(requestId, items, existing);
  const fileIdByUploadIndex = {};
  for (let i = 0; i < plan.length; i++) {
    const step = plan[i];
    let file = step.reuse ? findNamedImageFile_(folder, step.name) : null;
    if (!file) {
      const bytes = Utilities.base64Decode(uploads[step.index].data);
      const blobName = step.name || ("photo-" + step.index + ".jpg");
      file = folder.createFile(Utilities.newBlob(bytes, "image/jpeg", blobName));
      if (step.name) file.setName(step.name);
    }
    // 只開放這一張檔案。不要對資料夾 setSharing，避免整包照片被列出來。
    try {
      file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    } catch (err) {
      Logger.log("參考圖連結分享失敗：" + (err && err.message ? err.message : err));
    }
    fileIdByUploadIndex[step.index] = file.getId();
  }
  return collectImageLinks_(items, fileIdByUploadIndex);
}

function trashRemovedImages_(previousText, nextText) {
  const prev = parseImageCell_(previousText);
  if (!prev.length) return;
  const nextMap = {};
  const kept = parseImageCell_(nextText);
  for (let i = 0; i < kept.length; i++) nextMap[kept[i]] = true;
  let removed = false;
  for (let i = 0; i < prev.length; i++) if (!nextMap[prev[i]]) removed = true;
  if (!removed) return;

  const folderId = getOrCreateImageFolder_().getId();
  for (let i = 0; i < prev.length; i++) {
    if (nextMap[prev[i]]) continue;
    const id = imageFileId_(prev[i]);
    if (!id) continue;
    try {
      const file = DriveApp.getFileById(id);
      const parents = file.getParents();
      let ours = false;
      while (parents.hasNext()) {
        if (parents.next().getId() === folderId) ours = true;
      }
      if (ours && !file.isTrashed()) file.setTrashed(true);
    } catch (err) {
      Logger.log("移除參考圖失敗：" + (err && err.message ? err.message : err));
    }
  }
}

// 部署後在編輯器執行一次，同意雲端硬碟權限，並建立參考圖資料夾。
function authorizeDrive() {
  const folder = getOrCreateImageFolder_();
  Logger.log("已可使用雲端硬碟。參考圖資料夾：「" + folder.getName() + "」");
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
  const lastRow = Math.max(sheet.getLastRow(), 1);
  const width = Math.max(IMAGE_COL, sheet.getLastColumn());
  // 讀到 O 欄，參考圖片連結才會跟著寫進日曆說明。
  const data  = sheet.getRange(1, 1, lastRow, width).getValues();
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
      imageLinks: parseImageCell_(row[IMAGE_COL - 1]),
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

function imageLinesForEvent_(order) {
  const raw = order && (order.imageLinks != null ? order.imageLinks : order.images);
  const parts = Array.isArray(raw) ? raw : String(raw || '').split(/\s+/);
  const lines = [];
  const seen = {};
  for (let i = 0; i < parts.length; i++) {
    const s = String(parts[i] || '').trim();
    if (!/^https:\/\//i.test(s) || seen[s]) continue;
    seen[s] = true;
    lines.push(s);
  }
  return lines;
}

function orderEventDescription_(order) {
  const lines = ['客戶：' + (order.name || '')];
  if (order.phone) lines.push('電話：' + order.phone);
  if (order.amount) lines.push('金額：$' + order.amount);
  if (order.delivery) lines.push('配送：' + order.delivery);
  if (order.item) lines.push('品項：' + order.item);
  if (order.pickupInfo) lines.push('收花資訊：' + order.pickupInfo);
  const images = imageLinesForEvent_(order);
  if (images.length) {
    lines.push('參考圖片：');
    for (let i = 0; i < images.length; i++) lines.push(images[i]);
  }
  // 訂單編號維持最後一行，日曆對帳才找得到。
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