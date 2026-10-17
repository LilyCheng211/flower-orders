// 日曆對帳測試。執行：node tests/calendar-sync.test.js
const fs = require('fs');
const path = require('path');
const vm = require('vm');

const src = fs.readFileSync(path.join(__dirname, '..', 'Appscript.ts'), 'utf8');
const block = src.match(/\/\*__CALENDAR_SYNC_START__\*\/([\s\S]*?)\/\*__CALENDAR_SYNC_END__\*\//);
if (!block) {
  console.error('找不到日曆對帳程式');
  process.exit(1);
}
const sandbox = {};
vm.createContext(sandbox);
vm.runInContext(block[1], sandbox);

const planOrderCalendarSync_ = sandbox.planOrderCalendarSync_;
const orderEventContent_ = sandbox.orderEventContent_;
const orderEventTitle_ = sandbox.orderEventTitle_;
const orderIdFromDescription_ = sandbox.orderIdFromDescription_;
const isLegacyOrderEvent_ = sandbox.isLegacyOrderEvent_;
const isStallEventTitle_ = sandbox.isStallEventTitle_;

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
function idList(list) {
  return list.slice().sort().join(',');
}

function baseOrder(over) {
  return Object.assign({
    orderId: 'oaaaabbbbccccdddd',
    name: '王小明',
    phone: '0912345678',
    amount: 1500,
    delivery: '自取',
    item: '《 小花束 》粉色',
    pickupInfo: '',
    pickupTime: '15:00',
    day: '2026-10-18',
  }, over || {});
}

function eventFromContent(id, content, over) {
  return Object.assign({
    id: id,
    title: content.title,
    description: content.description,
    day: content.day,
    allDay: content.allDay,
    startH: content.startH,
    startM: content.startM,
    endH: content.endH,
    endM: content.endM,
  }, over || {});
}

// 舊程式建立的事件：標題一樣，但說明裡沒有訂單編號
function legacyEvent(id, order, over) {
  const content = orderEventContent_(order);
  const description = content.description.split('\n').filter(function (line) {
    return line.indexOf('【花店訂單：') !== 0;
  }).join('\n');
  return eventFromContent(id, content, Object.assign({ description: description }, over || {}));
}

function updateIds(plan) {
  return idList(plan.update.map(function (u) { return u.eventId; }));
}
function createIds(plan) {
  return idList(plan.create.map(function (o) { return o.orderId; }));
}
function assertNoOverlap(plan, msg) {
  const removed = {};
  let dup = false;
  plan.remove.forEach(function (id) {
    if (removed[id]) dup = true;
    removed[id] = true;
  });
  let overlap = false;
  plan.update.forEach(function (u) {
    if (removed[u.eventId]) overlap = true;
  });
  assert(!dup && !overlap, msg + '：更新和刪除沒有撞到同一筆');
}

// 把計畫套到假事件上，再對一次。第二次不該再新增或刪除。
function applyPlan(events, plan) {
  const removed = {};
  plan.remove.forEach(function (id) { removed[id] = true; });
  const updated = {};
  plan.update.forEach(function (u) { updated[u.eventId] = u.order; });
  const next = [];
  events.forEach(function (ev) {
    if (removed[ev.id]) return;
    if (updated[ev.id]) next.push(eventFromContent(ev.id, orderEventContent_(updated[ev.id])));
    else next.push(ev);
  });
  plan.create.forEach(function (order) {
    next.push(eventFromContent('new-' + order.orderId, orderEventContent_(order)));
  });
  return next;
}
function assertStable(orders, events, msg) {
  const plan = planOrderCalendarSync_(orders, events);
  assertNoOverlap(plan, msg);
  const again = planOrderCalendarSync_(orders, applyPlan(events, plan));
  eq(again.create.length, 0, msg + '：再跑一次不新增');
  eq(again.update.length, 0, msg + '：再跑一次不更新');
  eq(again.remove.length, 0, msg + '：再跑一次不刪除');
  eq(again.skip, orders.length, msg + '：再跑一次全部已是最新');
  return plan;
}

eq(orderEventTitle_('王小明', '《 小花束 》粉色'), '🌸 王小明｜小花束粉色', '標題去掉書名號');
eq(orderIdFromDescription_('客戶：王小明\n【花店訂單：oabc1234deadbeef】'), 'oabc1234deadbeef', '讀得出訂單編號');
eq(orderIdFromDescription_('客戶：王小明\n收花資訊：見【花店訂單：notthis】\n【花店訂單：yes12345】'), 'yes12345', '編號必須單獨一行');
eq(orderIdFromDescription_('客戶：王小明\n收花資訊：見【花店訂單：notthis】'), '', '句子裡的編號不當成鑰匙');
assert(isStallEventTitle_('🎪 出攤｜台科大') === true, '出攤標題');
assert(isLegacyOrderEvent_('🌸 王小明｜小花束', '客戶：王小明\n金額：$100') === true, '舊訂單事件');
assert(isLegacyOrderEvent_('🌸 生日｜提醒', '自己記的，不是訂單') === false, '手寫的花標記不算系統事件');
assert(isLegacyOrderEvent_('看醫生', '') === false, '一般行程不算系統事件');

const stall = {
  id: 'stall', title: '🎪 出攤｜台科大', description: '出攤地點：台科大',
  day: '2026-10-18', allDay: true, startH: null, startM: null, endH: null, endM: null,
};
const doctor = {
  id: 'doc', title: '看醫生', description: '自己的行程',
  day: '2026-10-18', allDay: false, startH: 9, startM: 0, endH: 10, endM: 0,
};
const handwritten = {
  id: 'note', title: '🌸 生日｜提醒', description: '自己記的，不是訂單',
  day: '2026-10-18', allDay: true, startH: null, startM: null, endH: null, endM: null,
};

// 1. 取花日期 10/17 → 10/18，新的那天還沒有事件：把舊的搬走，不要再新增一筆
(function () {
  const order = baseOrder({ day: '2026-10-18' });
  const old = legacyEvent('old', order, { day: '2026-10-17' });
  const plan = assertStable([order], [old, stall, doctor], '只改日期');
  eq(updateIds(plan), 'old', '只改日期：更新舊事件');
  eq(plan.create.length, 0, '只改日期：不新增');
  eq(plan.remove.length, 0, '只改日期：不刪除');
})();

// 2. 舊同步已經在新日期又建了一筆，10/17 還留著：留下 10/18，刪掉 10/17
(function () {
  const order = baseOrder({ day: '2026-10-18' });
  const old = legacyEvent('old', order, { day: '2026-10-17' });
  const neu = legacyEvent('neu', order);
  const plan = assertStable([order], [old, neu, stall, handwritten], '改日期後兩邊都有');
  eq(updateIds(plan), 'neu', '改日期後兩邊都有：留在新日期');
  eq(idList(plan.remove), 'old', '改日期後兩邊都有：刪掉舊日期');
  eq(plan.create.length, 0, '改日期後兩邊都有：不新增');
})();

// 3. 已經有編號，人在 10/17、訂單改成 10/18：還是改同一筆
(function () {
  const order = baseOrder({ day: '2026-10-18' });
  const ev = eventFromContent('e1', orderEventContent_(order), { day: '2026-10-17' });
  const plan = assertStable([order], [ev], '編號還在但日期是舊的');
  eq(updateIds(plan), 'e1', '編號還在但日期是舊的：更新同一筆');
  eq(plan.create.length, 0, '編號還在但日期是舊的：不新增');
  eq(plan.remove.length, 0, '編號還在但日期是舊的：不刪除');
})();

// 4. 改時間、金額、品項：不要留下舊內容的另一筆
(function () {
  const order = baseOrder({ pickupTime: '18:00', amount: 1800 });
  const ev = eventFromContent('e1', orderEventContent_(baseOrder({ pickupTime: '15:00', amount: 1500 })));
  const plan = assertStable([order], [ev], '改時間和金額');
  eq(updateIds(plan), 'e1', '改時間和金額：更新同一筆');
  eq(plan.create.length, 0, '改時間和金額：不新增');
  eq(plan.remove.length, 0, '改時間和金額：不刪除');
})();

(function () {
  const order = baseOrder({ item: '《 玫瑰 》紅' });
  const oldTitle = legacyEvent('old', baseOrder({ item: '《 小花束 》粉色' }));
  const neu = legacyEvent('neu', order);
  const plan = assertStable([order], [oldTitle, neu], '改品項後新舊標題都在');
  eq(updateIds(plan), 'neu', '改品項：留新標題');
  eq(idList(plan.remove), 'old', '改品項：刪舊標題');
  eq(plan.create.length, 0, '改品項：不新增');
})();

// 5. 刪掉的訂單，舊事件要消失；出攤和自己的行程留著
(function () {
  const ghost = legacyEvent('ghost', baseOrder({ name: '已刪除', orderId: 'o0000111122223333' }));
  const plan = assertStable([], [ghost, stall, doctor, handwritten], '刪除舊式訂單');
  eq(idList(plan.remove), 'ghost', '刪除舊式訂單：只刪那筆');
  eq(plan.create.length, 0, '刪除舊式訂單：不新增');
})();

(function () {
  const ghostOrder = baseOrder({ name: '已刪除', orderId: 'o0000111122223333' });
  const ghost = eventFromContent('ghost', orderEventContent_(ghostOrder));
  const plan = assertStable([], [ghost, stall], '刪除已編號訂單');
  eq(idList(plan.remove), 'ghost', '刪除已編號訂單：刪掉事件');
})();

// 6. 兩張不同的單，各留各的，不會互相刪
(function () {
  const a = baseOrder({ orderId: 'oaaaaaaaaaaaaaaa', name: '王小明', day: '2026-10-18' });
  const b = baseOrder({ orderId: 'obbbbbbbbbbbbbbb', name: '陳美美', day: '2026-10-19', pickupTime: '10:00', item: '玫瑰' });
  const events = [eventFromContent('ea', orderEventContent_(a)), eventFromContent('eb', orderEventContent_(b)), stall];
  const plan = assertStable([a, b], events, '兩張不同訂單');
  eq(plan.skip, 2, '兩張不同訂單：都已是最新');
  eq(plan.remove.length, 0, '兩張不同訂單：不刪除');
})();

// 7. 兩張同名同品項，日期不同。舊的 10/17 是改期留下的，不能把另一張單的事件刪掉。
(function () {
  const a = baseOrder({ orderId: 'oaaaaaaaaaaaaaaa', day: '2026-10-18' });
  const b = baseOrder({ orderId: 'obbbbbbbbbbbbbbb', day: '2026-10-19' });
  const events = [
    legacyEvent('old', a, { day: '2026-10-17' }),
    legacyEvent('d18', a),
    legacyEvent('d19', b),
    doctor,
  ];
  const plan = assertStable([a, b], events, '同標題兩張單');
  eq(updateIds(plan), 'd18,d19', '同標題兩張單：各留自己那天');
  eq(idList(plan.remove), 'old', '同標題兩張單：只刪多出來的舊日期');
  eq(plan.create.length, 0, '同標題兩張單：不新增');
})();

// 8. 完全沒有事件就新增；已經最新就略過
(function () {
  const order = baseOrder();
  const plan = assertStable([order], [stall, doctor], '全新訂單');
  eq(createIds(plan), order.orderId, '全新訂單：新增一筆');
  eq(plan.remove.length, 0, '全新訂單：不刪除出攤或自己的行程');
})();

(function () {
  const order = baseOrder();
  const ev = eventFromContent('e1', orderEventContent_(order));
  ev.description = ev.description.replace(/\n/g, '\r\n') + '\n';
  const plan = planOrderCalendarSync_([order], [ev]);
  eq(plan.skip, 1, '說明換行不同也算同一筆');
  eq(plan.update.length, 0, '說明換行不同：不重複更新');
  eq(plan.create.length, 0, '說明換行不同：不新增');
})();

// 9. 沒填時間是全日；後來補上 15:00，還是同一筆
(function () {
  const order = baseOrder({ pickupTime: '15:00' });
  const ev = eventFromContent('e1', orderEventContent_(baseOrder({ pickupTime: '寄出' })));
  const plan = assertStable([order], [ev], '全日改成有時間');
  eq(updateIds(plan), 'e1', '全日改成有時間：更新同一筆');
  eq(plan.create.length, 0, '全日改成有時間：不新增');
  assert(orderEventContent_(baseOrder({ pickupTime: '寄出' })).allDay === true, '寄出是全日');
  assert(orderEventContent_(order).allDay === false, '15:00 不是全日');
  eq(orderEventContent_(order).endM, 30, '15:00 結束在 15:30');
})();

// 10. 同一天重複兩筆舊事件，只留一筆
(function () {
  const order = baseOrder();
  const a = legacyEvent('a', order);
  const b = legacyEvent('b', order);
  const plan = assertStable([order], [a, b], '同一天重複');
  eq(plan.update.length, 1, '同一天重複：留一筆');
  eq(plan.remove.length, 1, '同一天重複：刪一筆');
  eq(plan.create.length, 0, '同一天重複：不新增');
  assert(plan.update[0].eventId !== plan.remove[0], '同一天重複：留下的和刪掉的不同');
})();

// 11. 試算表還在的舊訂單，行程留著；已經刪掉的舊訂單，殘留事件清掉
(function () {
  const may = legacyEvent('may', baseOrder({ name: '舊客戶', item: '花', day: '2026-05-06', orderId: 'o1111111111111111' }));
  const keep = [{ name: '舊客戶', title: orderEventTitle_('舊客戶', '花'), day: '2026-05-06' }];
  const active = baseOrder({ name: '舊客戶', item: '花', day: '2026-10-18', orderId: 'o2222222222222222' });
  const plan = planOrderCalendarSync_([active], [may, stall], keep, []);
  eq(createIds(plan), active.orderId, '同名舊行程不被新訂單搬走');
  eq(plan.update.length, 0, '同名舊行程：不改那筆舊的');
  eq(plan.remove.length, 0, '同名舊行程：不刪');
})();

(function () {
  const may = legacyEvent('may', baseOrder({ name: '已刪除', day: '2026-05-06' }));
  const plan = planOrderCalendarSync_([], [may, stall, doctor], [], []);
  eq(idList(plan.remove), 'may', '刪掉的舊訂單：殘留事件清掉');
})();

(function () {
  const oldOrder = baseOrder({ orderId: 'o3333333333333333', day: '2026-05-06' });
  const ev = eventFromContent('e1', orderEventContent_(oldOrder));
  const plan = planOrderCalendarSync_([], [ev], [], [oldOrder.orderId]);
  eq(plan.remove.length, 0, '已有編號的舊訂單還在，事件不刪');
  eq(plan.create.length, 0, '已有編號的舊訂單：不補建');
})();

(function () {
  const active = baseOrder({ day: '2026-10-18' });
  const may = legacyEvent('may', active, { day: '2026-05-06' });
  const plan = assertStable([active], [may], '很早的日期改到這個月');
  eq(updateIds(plan), 'may', '很早的日期改到這個月：把舊事件搬走');
  eq(plan.create.length, 0, '很早的日期改到這個月：不新增');
  eq(plan.remove.length, 0, '很早的日期改到這個月：不留下舊的');
})();

if (failed) {
  console.error('共 ' + failed + ' 項失敗');
  process.exit(1);
}
console.log('全部通過');
