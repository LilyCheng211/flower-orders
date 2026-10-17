// 訂單貼上解析測試。執行：node tests/parse-order.test.js
const fs = require('fs');
const path = require('path');
const vm = require('vm');

const html = fs.readFileSync(path.join(__dirname, '..', 'index.html'), 'utf8');
const block = html.match(/\/\*__PARSER_START__\*\/([\s\S]*?)\/\*__PARSER_END__\*\//);
if (!block) {
  console.error('找不到解析程式');
  process.exit(1);
}
const sandbox = {};
vm.createContext(sandbox);
vm.runInContext(block[1], sandbox);
const parse = sandbox.parseOrderMessage;
const norm = sandbox.normalizePickupTime;
const fixPhone = sandbox.restoreLeadingZero;

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

const template = [
  '請給我 >>>',
  '➤ LINE Pay 交易代碼：L1234',
  '➤ 訂購人姓名：王小明',
  '➤ 聯絡電話：0912345678',
  '➤ 取花日期：11/20',
  '➤ 幾點來取花：下午31點半',
  '(填一個確切時間，避免順序錯亂，敬請準時到達，勿提早與遲到)'
].join('\n');
const a = parse(template);
eq(a.last5, 'L1234', '範本：交易代碼');
eq(a.name, '王小明', '範本：姓名');
eq(a.phone, '0912345678', '範本：電話');
eq(a.pickup_date, '11/20', '範本：日期');
eq(a.pickup_time, '15:30', '範本：下午31點半 → 15:30');
eq(a.info, '', '範本：沒有收花資訊');
assert(a.info.indexOf('0912345678') < 0, '範本：收花資訊不含訂購人電話');

const oneLine = 'LINE Pay 交易代碼：A5678 / 陳美美 / 0987654321 / 12/24 19:00';
const b = parse(oneLine);
eq(b.last5, 'A5678', '一行：交易代碼');
eq(b.name, '陳美美', '一行：姓名');
eq(b.phone, '0987654321', '一行：電話');
eq(b.pickup_date, '12/24', '一行：日期');
eq(b.pickup_time, '19:00', '一行：時間');
eq(b.info, '', '一行：沒有收花資訊');

const colloquial = [
  '交易代碼 L7918',
  'Jeffery',
  '0972915271',
  '11/9 晚上6點',
  '後山埤站全家新城鑫店'
].join('\n');
const c = parse(colloquial);
eq(c.last5, 'L7918', '口語：交易代碼');
eq(c.name, 'Jeffery', '口語：姓名');
eq(c.phone, '0972915271', '口語：電話');
eq(c.pickup_date, '11/9', '口語：日期');
eq(c.pickup_time, '18:00', '口語：晚上6點 → 18:00');
eq(c.note, '後山埤站全家新城鑫店', '口語：店名進備註');
eq(c.info, '', '口語：店名不進收花資訊');
assert(c.info.indexOf('0972915271') < 0, '口語：收花資訊不含訂購人電話');
eq(c.delivery, '全家 自取', '口語：看到全家，配送猜成全家自取');

const colloquialLine = 'L7918 Jeffery 0972915271 11/9 晚上6點 後山埤站全家新城鑫店';
const c2 = parse(colloquialLine);
eq(c2.last5, 'L7918', '同一行口語：交易代碼');
eq(c2.name, 'Jeffery', '同一行口語：姓名');
eq(c2.phone, '0972915271', '同一行口語：電話');
eq(c2.pickup_date, '11/9', '同一行口語：日期');
eq(c2.pickup_time, '18:00', '同一行口語：晚上6點 → 18:00');
assert(c2.note.indexOf('後山埤站全家新城鑫店') >= 0, '同一行口語：店名進備註');
eq(c2.info, '', '同一行口語：不進收花資訊');

const bouquet = [
  '《 求婚花束 》中型｜玫瑰為主',
  '自然風・白／粉／鵝黃色配色',
  '透明包裝＋雷射紙底＋marry me',
  '➤ 金額：5000'
].join('\n');
const d = parse(bouquet);
eq(d.item, '《求婚花束》中型｜玫瑰為主，自然風・白／粉／鵝黃色配色，透明包裝＋雷射紙底＋marry me', '單一花束整理成一行');
eq(d.amount, '5000', '單一花束金額');
eq(d.name, '', '花束英文 me 不當成姓名');

const two = [
  '《 求婚花束 》中型｜玫瑰為主',
  '自然風・白／粉／鵝黃色配色',
  '➤ 金額：5000',
  '',
  '《 小花束 》粉色系',
  '➤ 金額：800'
].join('\n');
const e = parse(two);
eq(e.item, '《求婚花束》中型｜玫瑰為主，自然風・白／粉／鵝黃色配色；《小花束》粉色系', '多個花束用分號接起來');
eq(e.amount, '5800', '多個花束金額加總');

const recipient = [
  '➤ 訂購人姓名：王小明',
  '➤ 聯絡電話：0912345678',
  '➤ 取花日期：11/9',
  '➤ 幾點來取花：上午10點',
  '收花人：林小華',
  '收花人電話：0911000222',
  '配送地址：台北市大安區復興南路1號'
].join('\n');
const f = parse(recipient);
eq(f.name, '王小明', '收花人範例：訂購人');
eq(f.phone, '0912345678', '收花人範例：訂購人電話');
eq(f.pickup_time, '10:00', '收花人範例：上午10點');
assert(f.info.indexOf('林小華') >= 0, '收花資訊有收花人');
assert(f.info.indexOf('0911000222') >= 0, '收花資訊有收花人電話');
assert(f.info.indexOf('復興南路') >= 0, '收花資訊有地址');
assert(f.info.indexOf('0912345678') < 0, '收花資訊不含訂購人電話');

eq(norm('早上9點15分'), '09:15', '早上9點15分');
eq(norm('上午10點'), '10:00', '上午10點');
eq(norm('中午12點半'), '12:30', '中午12點半');
eq(norm('下午3點半'), '15:30', '下午3點半');
eq(norm('下午31點半'), '15:30', '下午31點半');
eq(norm('傍晚5點'), '17:00', '傍晚5點');
eq(norm('晚上6點'), '18:00', '晚上6點');
eq(norm('19:00'), '19:00', '19:00');
eq(norm('下午1點20分'), '13:20', '下午1點20分');

eq(fixPhone('972915271'), '0972915271', '九位數補 0');
eq(fixPhone('0972915271'), '0972915271', '已有 0 不重複補');
eq(fixPhone('972915271.0'), '0972915271', '小數形式補 0');
eq(fixPhone(''), '', '空電話');
eq(fixPhone('12345'), '12345', '不是手機就不補 0');

assert(sandbox.deliveryNeedsInfo('黑貓 上午') === true, '黑貓需要收花資訊');
assert(sandbox.deliveryNeedsInfo('黑貓 下午') === true, '黑貓下午需要收花資訊');
assert(sandbox.deliveryNeedsInfo('Lala 來取') === true, 'Lala 需要收花資訊');
assert(sandbox.deliveryNeedsInfo('Lala 代叫') === true, 'Lala 代叫需要收花資訊');
assert(sandbox.deliveryNeedsInfo('親自外送') === true, '親自外送需要收花資訊');
assert(sandbox.deliveryNeedsInfo('全家 自取') === false, '全家自取不強調收花資訊');
assert(sandbox.deliveryNeedsInfo('自取') === false, '自取不強調收花資訊');
assert(sandbox.deliveryNeedsInfo('TBC') === false, 'TBC 不強調收花資訊');

const oldLine = '李品諺 0952288391 64364 5/7 18:00';
const g = parse(oldLine);
eq(g.name, '李品諺', '舊一行格式：姓名');
eq(g.phone, '0952288391', '舊一行格式：電話');
eq(g.last5, '64364', '舊一行格式：五碼仍可讀');
eq(g.pickup_date, '5/7', '舊一行格式：日期');
eq(g.pickup_time, '18:00', '舊一行格式：時間');
eq(g.info, '', '舊一行格式：不把電話放進收花資訊');

if (failed) {
  console.error('共 ' + failed + ' 項失敗');
  process.exit(1);
}
console.log('全部通過');
