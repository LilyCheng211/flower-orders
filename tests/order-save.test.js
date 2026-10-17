// 訂單儲存、備註與重複送出測試。執行：node tests/order-save.test.js
const fs = require('fs');
const path = require('path');
const vm = require('vm');

function loadBlock(file, start, end) {
  const src = fs.readFileSync(path.join(__dirname, '..', file), 'utf8');
  const match = src.match(new RegExp(start + '([\\s\\S]*?)' + end));
  if (!match) {
    console.error('找不到程式區塊：' + start);
    process.exit(1);
  }
  return match[1];
}

const sandbox = {};
vm.createContext(sandbox);
vm.runInContext(loadBlock('index.html', '\\/\\*__PARSER_START__\\*\\/', '\\/\\*__PARSER_END__\\*\\/'), sandbox);
vm.runInContext(loadBlock('index.html', '\\/\\*__ORDER_CLIENT_START__\\*\\/', '\\/\\*__ORDER_CLIENT_END__\\*\\/'), sandbox);
vm.runInContext(loadBlock('Appscript.ts', '\\/\\*__ORDER_IO_START__\\*\\/', '\\/\\*__ORDER_IO_END__\\*\\/'), sandbox);

let failed = 0;
function assert(cond, msg) {
  if (!cond) {
    failed++;
    console.error('失敗：' + msg);
  } else {
    console.log('通過：' + msg);
  }
}
function eq(actual, expected, msg) {
  assert(actual === expected, msg + '｜實際：' + JSON.stringify(actual) + '｜預期：' + JSON.stringify(expected));
}

const note = '許願色系：粉橘；\r\n不要滿天星\n請附卡片';
const cleanNote = '許願色系：粉橘；\n不要滿天星\n請附卡片';
eq(sandbox.normalizeNoteText(note), cleanNote, '前端備註保留換行與；');
eq(sandbox.normalizeNote_('  ' + note + '  '), cleanNote, '後端備註保留換行與；');

const data = {
  name: '水晶',
  phone: '0912345678',
  amount: '5000',
  pickup_date: '10/23',
  pickup_time: '09:00',
  delivery: '自取',
  item: '水晶花束',
  last5: 'L5000',
  pickup_info: '收花人：林小華\n地址：復興南路',
  card_text: '生日快樂\n第二行',
  note: cleanNote
};
const cells = sandbox.orderToCells_(data, false);
eq(cells.length, 12, '一列寫入 12 欄，備註在最後');
eq(cells[11], cleanNote, '備註寫在 L 欄（第 12 格）');
eq(cells[0], '水晶', '姓名仍在 A 欄');
eq(cells[9], '收花人：林小華\n地址：復興南路', '收花資訊的換行也留著');

const back = sandbox.orderFromCells_(cells, 8);
eq(back.note, cleanNote, '讀回備註，換行與；都在');
eq(back.name, '水晶', '讀回姓名');
eq(back.row, 8, '列號');
assert(back.paid === false, '新訂單尚未收款');

const seen = ['rabc123', 'rother99'];
const again = sandbox.decideDuplicate_(seen, 'rabc123');
assert(again.duplicate === true && again.row === 2, '同一個請求編號不新增第二列');
const fresh = sandbox.decideDuplicate_(seen, 'rnew999');
assert(fresh.duplicate === false && fresh.row === 0, '新的請求編號可以新增');
const bad = sandbox.decideDuplicate_(seen, '不是編號');
assert(bad.duplicate === false && bad.id === '', '不像編號的值不當成重複鍵');

const body = sandbox.encodeWriteBody({
  action: 'addOrder',
  requestId: 'rnote123',
  data: { note: cleanNote, name: '水晶' }
});
assert(body.indexOf('\n') < 0, '送出的 JSON 本文沒有裸換行');
const parsed = sandbox.parseWriteBody(body);
eq(parsed.data.note, cleanNote, 'POST 本文讀回多行備註');
eq(parsed.data.name, '水晶', 'POST 本文讀回中文姓名');

// 若把備註直接放進網址、遇到換行就被截斷，JSON 會壞掉。所以改走整份本文。
const cut = '{"name":"水晶","note":"許願色系：粉橘；'.split('\n')[0];
let cutThrew = false;
try { JSON.parse(cut); } catch (e) { cutThrew = true; }
assert(cutThrew, '網址在換行被截斷時，整包 JSON 不能解析');

let seq = 0;
function fakeId() { seq += 1; return 'r' + String(seq).padStart(6, '0'); }
const first = sandbox.sameSubmissionId({ id: '', fingerprint: '' }, 'same', fakeId);
const second = sandbox.sameSubmissionId(first, 'same', fakeId);
eq(second.id, first.id, '同樣內容再按一次，沿用同一個請求編號');
const third = sandbox.sameSubmissionId(second, 'changed', fakeId);
assert(third.id !== second.id, '內容改了就換新的請求編號');

const lines = sandbox.orderDetailLines({
  name: '水晶',
  phone: 912345678,
  amount: 5000,
  pickup_date: '10/23',
  pickup_time: '09:00',
  delivery: '自取',
  item: '水晶花束',
  last5: 'L5000',
  paid: true,
  pickup_info: '',
  card_text: '生日快樂',
  note: cleanNote
});
function line(label) {
  const found = lines.filter(function(item){ return item.label === label; })[0];
  return found ? found.value : null;
}
eq(lines.length, 12, '明細有 12 個欄位');
eq(line('姓名'), '水晶', '明細有姓名');
eq(line('電話'), '0912345678', '明細電話補回開頭的 0');
eq(line('金額'), '$5,000', '明細金額');
eq(line('取花日期'), '10/23', '明細日期');
eq(line('取花時間'), '09:00', '明細時間');
eq(line('配送'), '自取', '明細配送');
eq(line('品項'), '水晶花束', '明細品項');
eq(line('LINE Pay 交易代碼'), 'L5000', '明細交易代碼');
eq(line('收到款'), '已收款', '明細收款');
eq(line('備註'), cleanNote, '明細備註保留換行與；');

const kept = sandbox.nextOrdersAfterRefresh(
  [{ name: '水晶', note: cleanNote }],
  { ok: false }
);
eq(kept.length, 1, '重新整理失敗時留下原本的訂單');
eq(kept[0].note, cleanNote, '失敗時備註還在');
const replaced = sandbox.nextOrdersAfterRefresh(
  [{ name: '舊的' }],
  { ok: true, orders: [{ name: '新的' }] }
);
eq(replaced[0].name, '新的', '重新整理成功才換成新資料');

assert(sandbox.shouldStartOrderSync(true, false) === true, '已解鎖就要開始抓訂單');
assert(sandbox.shouldStartOrderSync(false, true) === true, '示範模式一進頁就抓訂單');
assert(sandbox.shouldStartOrderSync(false, false) === false, '還沒解鎖不要抓、也不要畫成 0 張');

const html = fs.readFileSync(path.join(__dirname, '..', 'index.html'), 'utf8');
assert(html.indexOf("method: 'POST'") >= 0, '寫入使用 POST');
assert(html.indexOf('text/plain;charset=utf-8') >= 0, '寫入用純文字本文，避免預檢被擋');
assert(html.indexOf('AbortController') >= 0, '連線有逾時');
const refreshAt = html.indexOf('async function manualRefresh');
const refreshFn = html.slice(refreshAt, html.indexOf('function startAutoRefresh'));
assert(refreshFn.indexOf('allOrders=[]') < 0 && refreshFn.indexOf('allOrders = []') < 0, '手動刷新不會先把訂單清成空的');
assert(refreshFn.indexOf('finally') >= 0, '手動刷新會把按鈕還原');
assert(html.indexOf('for (let tryAgain') < 0 && html.indexOf('apiWrite(') >= 0, '有寫入函式');
const writeFn = html.slice(html.indexOf('function apiWrite'), html.indexOf('function demoWrite'));
const writeBody = writeFn.slice(writeFn.indexOf('{'));
assert(writeBody.indexOf('apiWrite') < 0, '寫入函式自己不會再呼叫一次');

const unlockFn = html.slice(html.indexOf('async function unlock'), html.indexOf('if(sessionStorage.getItem'));
assert(unlockFn.indexOf('startOrderSync()') >= 0, '解鎖成功後立刻抓訂單');
const boot = html.slice(html.lastIndexOf('document.body.addEventListener'));
assert(boot.indexOf('startOrderSync()') >= 0, '重新整理且已登入時立刻抓訂單');
assert(boot.indexOf('renderCalendar()') < 0 && boot.indexOf('renderStats()') < 0, '進頁時不要先用空清單畫月曆');
const autoFn = html.slice(html.indexOf('function startAutoRefresh'), html.indexOf('function renderStats'));
assert(autoFn.indexOf('loadOrders(true)') >= 0, '每 3 分鐘的自動同步會真的再抓一次');
const addFn = html.slice(html.indexOf('async function submitOrder'), html.indexOf('function renderAll'));
assert(addFn.indexOf("readOrderForm('f')") >= 0, '新增訂單會讀表單，含備註');
assert(addFn.indexOf("action: 'addOrder'") >= 0, '新增走 addOrder');
assert(addFn.indexOf('allOrders = []') < 0 && addFn.indexOf('allOrders=[]') < 0, '新增成功不會先把訂單清成空的');
assert(addFn.indexOf('loadOrders(true)') >= 0, '新增成功後會再抓一次，畫面上看得到新訂單');

const script = fs.readFileSync(path.join(__dirname, '..', 'Appscript.ts'), 'utf8');
const insertFn = script.slice(script.indexOf('function insertOrder_'), script.indexOf('function updateOrder'));
assert(insertFn.indexOf('orderToCells_') >= 0, '新增用 orderToCells_ 組列，備註在第 12 格');
assert(insertFn.indexOf('writeOrderRow_') >= 0, '新增會整列寫入');
assert(insertFn.indexOf('NOTE_COL') >= 0 && insertFn.indexOf('cells[11]') >= 0, '新增後再把備註寫進 L 欄');
const createBody = sandbox.encodeWriteBody({
  action: 'addOrder',
  requestId: 'rnote123',
  data: data
});
const createParsed = sandbox.parseWriteBody(createBody);
eq(createParsed.action, 'addOrder', '新增的本文是 addOrder');
eq(createParsed.data.note, cleanNote, '新增本文帶著多行備註與；');
eq(sandbox.orderToCells_(createParsed.data, false)[11], cleanNote, '新增本文的備註會進 L 欄');

if (failed) {
  console.error('共 ' + failed + ' 項失敗');
  process.exit(1);
}
console.log('全部通過');
