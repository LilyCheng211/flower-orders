// 參考圖片測試。執行：node tests/order-images.test.js
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
vm.runInContext(loadBlock('index.html', '\\/\\*__ORDER_CLIENT_START__\\*\\/', '\\/\\*__ORDER_CLIENT_END__\\*\\/'), sandbox);
vm.runInContext(loadBlock('Appscript.ts', '\\/\\*__ORDER_IO_START__\\*\\/', '\\/\\*__ORDER_IO_END__\\*\\/'), sandbox);
vm.runInContext(loadBlock('Appscript.ts', '\\/\\*__CALENDAR_SYNC_START__\\*\\/', '\\/\\*__CALENDAR_SYNC_END__\\*\\/'), sandbox);

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

const keepA = 'https://drive.google.com/file/d/keep123456/view';
const keepB = 'https://lh3.googleusercontent.com/d/keep654321';
const openB = 'https://drive.google.com/file/d/keep654321/view';

eq(sandbox.fitImageSize(4000, 3000, 1600).width, 1600, '手機橫圖長邊縮到 1600');
eq(sandbox.fitImageSize(4000, 3000, 1600).height, 1200, '手機橫圖高度跟著縮');
eq(sandbox.fitImageSize(3000, 4000).width, 1200, '手機直圖寬度跟著縮');
eq(sandbox.fitImageSize(3000, 4000).height, 1600, '手機直圖長邊縮到 1600');
eq(sandbox.fitImageSize(800, 600, 1600).width, 800, '小圖不放大');
eq(sandbox.fitImageSize(0, 10).width, 0, '沒有尺寸就不用縮');

eq(sandbox.approxBase64Bytes('YQ=='), 1, 'base64 長度換算位元組');
eq(sandbox.pickJpegQuality([
  { quality: 0.72, bytes: 800000 },
  { quality: 0.6, bytes: 400000 },
  { quality: 0.48, bytes: 200000 }
], 450 * 1024), 0.6, '挑第一個塞得進上限的品質');
eq(sandbox.pickJpegQuality([
  { quality: 0.72, bytes: 900000 },
  { quality: 0.48, bytes: 800000 }
], 450 * 1024), 0.48, '都太大就用最低的品質');

let drafts = [];
for (let i = 0; i < 6; i++) {
  const added = sandbox.addImageDraft(drafts, { data: 'YQ==', name: 'photo.jpg' }, 6);
  assert(added.ok === true, '第 ' + (i + 1) + ' 張可以加入');
  drafts = added.list;
}
const tooMany = sandbox.addImageDraft(drafts, { data: 'YQ==' }, 6);
assert(tooMany.ok === false && tooMany.list.length === 6, '最多 6 張');
const removed = sandbox.removeImageDraft(drafts, 0);
eq(removed.length, 5, '編輯時可以移除一張');
eq(sandbox.imagesToPayload(removed)[0].data, 'YQ==', '剩下的仍是要上傳的圖片');

const mixed = [
  { link: keepA },
  { data: 'YQ==', name: 'photo.jpg', preview: 'data:image/jpeg;base64,YQ==' }
];
const payload = sandbox.imagesToPayload(mixed);
eq(payload[0].url, keepA, '已上傳的圖只送連結');
eq(payload[1].data, 'YQ==', '新圖才送檔案內容');
eq(sandbox.demoImageLinks(payload)[1], 'data:image/jpeg;base64,YQ==', '示範模式用 data 網址顯示新圖');

eq(sandbox.imageFileId(keepA), 'keep123456', '從雲端硬碟檢視連結取出檔案編號');
eq(sandbox.imageDisplayUrl(keepA), 'https://lh3.googleusercontent.com/d/keep123456', '前端用可顯示的圖片網址');
eq(sandbox.imageFallbackUrl(keepA).indexOf('keep123456') >= 0, true, '顯示失敗還有縮圖網址');
eq(sandbox.orderImageLinks({ images: [keepA, '不是圖'] }).length, 1, '明細只留看得懂的圖片');
eq(sandbox.orderDetailLines({ name: '水晶', images: [keepA] }).length, 12, '有圖片時文字欄仍是 12 個');

const fpEmpty = sandbox.submissionFingerprint({ name: '水晶', images: [] });
const fpPhoto = sandbox.submissionFingerprint({ name: '水晶', images: [{ data: 'YQ==' }] });
assert(fpEmpty !== fpPhoto, '圖片變了，這次送出的內容就算不同');
const same = sandbox.sameSubmissionId({ id: '', fingerprint: '' }, fpPhoto, function () { return 'rphoto1'; });
const retry = sandbox.sameSubmissionId(same, fpPhoto, function () { return 'rphoto2'; });
eq(retry.id, 'rphoto1', '同樣的圖重送，沿用同一個請求編號');
const changed = sandbox.sameSubmissionId(retry, fpEmpty, function () { return 'rphoto3'; });
assert(changed.id !== retry.id, '拿掉圖片就換新的請求編號');

eq(sandbox.validateIncomingImages_(undefined), '', '沒有圖片不算錯誤');
eq(sandbox.validateIncomingImages_([]), '', '空的圖片清單可以送');
assert(sandbox.hasImageField_({ name: '水晶' }) === false, '舊頁面沒有 images 欄位');
assert(sandbox.hasImageField_({ images: [] }) === true, '新頁面有送 images');
const seven = [];
for (let i = 0; i < 7; i++) {
  seven.push({ url: 'https://drive.google.com/file/d/keep12345' + i + '/view' });
}
assert(sandbox.validateIncomingImages_(seven).indexOf('最多') >= 0, '超過 6 張會被拒絕');
assert(sandbox.validateIncomingImages_([{ data: 'A'.repeat(700001) }]).indexOf('太大') >= 0, '單張太大會被拒絕');
assert(sandbox.validateIncomingImages_([{ data: '###' }]).indexOf('格式') >= 0, '不是 base64 會被拒絕');

const items = sandbox.normalizeIncomingImages_([
  { url: keepA },
  { data: 'YQ==', name: 'a.jpg' },
  { url: keepB },
  { url: 'https://example.com/not-drive' }
]);
eq(items.length, 3, '不是雲端硬碟的連結不存');
eq(items[0].kind, 'keep', '原本的圖留著');
eq(items[1].kind, 'upload', '新圖要上傳');
eq(items[2].url, openB, '顯示用網址改存成可點開的檢視連結');

const planRetry = sandbox.planImageUploads_('rabc1234', items, ['ref-rabc1234-0.jpg']);
eq(planRetry.length, 1, '兩張舊圖中間只有一張新圖');
eq(planRetry[0].name, 'ref-rabc1234-0.jpg', '檔名綁請求編號');
assert(planRetry[0].reuse === true, '同名檔已經在資料夾就不要再上傳');
const planFresh = sandbox.planImageUploads_('rabc1234', items, []);
assert(planFresh[0].reuse === false, '第一次還沒有檔案');
const planAnon = sandbox.planImageUploads_('', items, ['ref-rabc1234-0.jpg']);
assert(planAnon[0].reuse === false && planAnon[0].name === '', '沒有請求編號不沿用別人的檔');

const saved = sandbox.collectImageLinks_(items, { 0: 'file999999' });
eq(saved, [keepA, 'https://drive.google.com/file/d/file999999/view', openB].join('\n'), '連結順序跟使用者排的一樣');

const plain = sandbox.orderToCells_({
  name: '水晶', phone: '0912345678', amount: '5000', pickup_date: '10/17', pickup_time: '09:00',
  delivery: '自取', item: '水晶花束', last5: 'L5000', pickup_info: '', card_text: '',
  note: '許願色系：粉橘；\n不要滿天星',
  images: [{ url: keepA }]
}, false);
eq(plain.length, 12, '圖片不寫進前 12 欄');
eq(plain[11], '許願色系：粉橘；\n不要滿天星', '備註仍在 L 欄');

const oldRow = ['水晶', '0912345678', '5000', '10/17', '09:00', '自取', '花', 'L5000', false, '', '', '只有備註'];
const oldOrder = sandbox.orderFromCells_(oldRow, 4);
eq(oldOrder.images.length, 0, '舊訂單沒有參考圖也讀得回來');
eq(oldOrder.note, '只有備註', '備註不會被當成圖片');

const row = oldRow.slice();
row[11] = '備註裡有連結 ' + keepA;
row[14] = keepA + '\nhttps://drive.google.com/open?id=keep654321\n不是連結';
const withImages = sandbox.orderFromCells_(row, 5);
eq(withImages.note, '備註裡有連結 ' + keepA, '備註裡的連結仍留在 L');
eq(withImages.images.length, 2, 'O 欄兩個連結都讀得到');
eq(withImages.images[1], openB, 'open?id= 也收成檢視連結');

const seen = ['rabc1234'];
const again = sandbox.decideDuplicate_(seen, 'rabc1234');
assert(again.duplicate === true && again.row === 2, '同一個請求編號不新增第二列');

const order = {
  orderId: 'oaaaabbbbccccdddd',
  name: '王小明',
  phone: '0912345678',
  amount: 1500,
  delivery: '自取',
  item: '《 小花束 》粉色',
  pickupInfo: '',
  pickupTime: '15:00',
  day: '2026-10-18',
  imageLinks: [keepA]
};
const described = sandbox.orderEventContent_(order).description;
assert(described.indexOf('參考圖片：') >= 0, '日曆說明有參考圖片');
assert(described.indexOf(keepA) >= 0, '日曆說明有連結');
eq(described.split('\n').pop(), '【花店訂單：oaaaabbbbccccdddd】', '訂單編號仍在最後一行');
eq(sandbox.orderIdFromDescription_(described), 'oaaaabbbbccccdddd', '有圖片仍讀得到訂單編號');

const without = sandbox.orderEventContent_(Object.assign({}, order, { imageLinks: [] }));
assert(without.description.indexOf('參考圖片') < 0, '沒有圖片的行程說明維持原樣');
const oldEvent = {
  id: 'ev1',
  title: without.title,
  description: without.description,
  day: without.day,
  allDay: without.allDay,
  startH: without.startH,
  startM: without.startM,
  endH: without.endH,
  endM: without.endM
};
const plan = sandbox.planOrderCalendarSync_([order], [oldEvent]);
eq(plan.create.length, 0, '補上參考圖不新增第二個行程');
eq(plan.update.length, 1, '補上參考圖只更新原來那筆行程');
eq(plan.remove.length, 0, '補上參考圖不刪掉原行程');

const script = fs.readFileSync(path.join(__dirname, '..', 'Appscript.ts'), 'utf8');
const html = fs.readFileSync(path.join(__dirname, '..', 'index.html'), 'utf8');
assert(script.indexOf('const IMAGE_COL = 15') >= 0, '參考圖片在第 15 欄 O');
assert(script.indexOf('花店訂單參考圖') >= 0, '圖片放在專用資料夾');
assert(script.indexOf('https://www.googleapis.com/auth/drive') >= 0, '註明雲端硬碟權限');
assert(script.indexOf('function authorizeDrive') >= 0, '有一次授權用的函式');
assert(script.indexOf('folder.setSharing') < 0, '不分享整個資料夾');
assert(script.indexOf('DriveApp.Access.ANYONE_WITH_LINK') >= 0, '單張圖知道連結的人可看');
assert(script.indexOf('DriveApp.Access.ANYONE)') < 0 && script.indexOf('DriveApp.Access.ANYONE,') < 0, '不是公開到整個網路');

const addFn = script.slice(script.indexOf('function addOrder'), script.indexOf('function insertOrder_'));
assert(addFn.indexOf('rememberedRow_') >= 0 && addFn.indexOf('rememberedRow_') < addFn.indexOf('insertOrder_'), '重複的請求在插入前就返回');
assert(addFn.indexOf('saveOrderImages_') < 0, '重複請求不會進到上傳');
const insertFn = script.slice(script.indexOf('function insertOrder_'), script.indexOf('function updateOrder'));
assert(insertFn.indexOf('saveOrderImages_') >= 0 && insertFn.indexOf('saveOrderImages_') < insertFn.indexOf('writeOrderRow_'), '先上傳圖片，成功才寫列');
assert(insertFn.indexOf('rememberRequestId_') > insertFn.indexOf('writeImageCell_'), '圖片寫進儲存格後才記住請求編號');
const saveFn = script.slice(script.indexOf('function saveOrderImages_'), script.indexOf('function trashRemovedImages_'));
assert(saveFn.indexOf('planImageUploads_') >= 0 && saveFn.indexOf('if (!file)') >= 0, '同名檔已在就不建立第二個');
assert(saveFn.indexOf('getOrCreateImageFolder_') > saveFn.indexOf('if (!uploads.length)'), '沒有新圖就不要打開雲端硬碟');

assert(html.indexOf('id="f-images"') >= 0 && html.indexOf('id="e-images"') >= 0, '新增和編輯都能選圖片');
assert(html.indexOf('accept="image/*"') >= 0, '可以選手機照片');
assert(html.indexOf('function compressImageFile') >= 0, '上傳前會縮小圖片');
assert(html.indexOf('data-act="remove-image"') >= 0, '編輯時可以移除縮圖');
assert(html.indexOf('data-act="zoom"') >= 0, '明細可以點開放大');
assert(script.indexOf('action === "login"') >= 0, '登入流程還在');
assert(script.indexOf('params.password === config.adminPassword') >= 0, '密碼仍由後端比對');

if (failed) {
  console.error('共 ' + failed + ' 項失敗');
  process.exit(1);
}
console.log('全部通過');
