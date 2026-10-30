// 振り分けの見本のテスト（偽物の Gmail とシート・AI なし）。node projects/orbit-gas-sample/test.js
const assert = require('assert');
const labels = {};
const mail = (件名, 本文) => ({ getDate: () => new Date('2026-10-25T09:00:00+09:00'), getSubject: () => 件名, getPlainBody: () => 本文 });
const inbox = [
  ['【試し】内見の問い合わせ', '東町の内見を11月1日にお願いしたいです'],
  ['【試し】内見の日程の変更', '内見の日にちを変えたいです'],
  ['【試し】キャンセルします', '先日の内見はキャンセルでお願いします'],
  ['【試し】家賃について', '家賃と初期費用を教えてください'],
  ['【試し】見学できますか', '週末に見学できますか'],
].map(([s, b]) => ({ msg: mail(s, b), added: [] }));
let 出口のスレッド = [];
let 受信 = () => inbox.map(t => ({ getMessages: () => [t.msg], getMessageCount: () => 1, addLabel: l => t.added.push(l.name) }));
global.GmailApp = {
  search: (q) => /label:orbit-返信待ち/.test(q) ? 出口のスレッド : 受信(q),
  getUserLabelByName: n => labels[n], createLabel: n => (labels[n] = { name: n }),
};
const props = { FIRE_URL: 'https://api.anthropic.com/v1/claude_code/routines/trig_x/fire', FIRE_TOKEN: 'tok' };
global.PropertiesService = { getScriptProperties: () => ({ getProperty: k => props[k] || null, setProperty: (k, v) => { props[k] = v; } }) };
const 表をつくる = tabs => ({ getId: () => 'x', getSheets: () => Object.values(tabs), getSheetByName: n => tabs[n] || null,
  insertSheet: n => (tabs[n] = タブ(n)) });
const タブ = n => { const t = { name: n, rows: [], appendRow: r => t.rows.push(r), getLastRow: () => t.rows.length }; return t; };
const 新しい表 = { シート1: タブ('シート1') }, 台帳 = { 仕事ごと: タブ('仕事ごと'), 回ごと: タブ('回ごと') };
let 開く = null;   // SHEET_ID が管理台帳を指す時の表
global.SpreadsheetApp = { create: () => 表をつくる(新しい表), openById: () => { if (!開く) throw new Error('無い'); return 開く; } };
const rows = 新しい表.シート1.rows;
const fired = [];
global.UrlFetchApp = { fetch: (url, o) => { if (!/\/fire$/.test(url)) throw new Error('/fire 以外に送ってはいけない'); fired.push({ url, o }); return { getResponseCode: () => 200, getContentText: () => '{"type":"routine_fire"}' }; } };
const g = require('./Code.js');

assert.equal(g.種類を決める('内見の問い合わせ', ''), '内見');
assert.equal(g.種類を決める('', '内見の日程を変えたい'), '変更');
assert.equal(g.種類を決める('', 'キャンセルで'), '取消');
assert.equal(g.種類を決める('家賃', '教えて'), '質問');
const n = g.振り分け();
assert.deepEqual(n, { 全部: 5, AI: 3, 呼んだ: true });
assert.equal(fired.length, 1, '返信待ちがある時だけ、まとめて1回呼ぶ');
const h = fired[0].o.headers, body = JSON.parse(fired[0].o.payload);
assert.ok(h.Authorization === 'Bearer tok' && h['anthropic-beta'] === 'experimental-cc-routine-2026-04-01', '公式のヘッダ');
assert.ok(body.text.includes('問い合わせ 3件') && body.text.split('---').length === 4, '返信待ちの3件だけを text に');
assert.ok(inbox.every(t => t.added.includes('orbit-済み')), '全部に済みのラベル');
assert.ok(inbox[1].added.includes('orbit-変更・取消') && inbox[2].added.includes('orbit-変更・取消'), '変更・取消は人へ');
assert.equal(rows.length, 1 + 5, '見出し＋5行');

// 管理台帳（「回ごと」のタブがある表）を SHEET_ID に入れた時：中に「問い合わせの記録」を作り、回ごとに1行だけ足す
props.SHEET_ID = 'ledger'; 開く = 表をつくる(台帳);
inbox.forEach(t => (t.added = []));
assert.deepEqual(g.振り分け(), { 全部: 5, AI: 3, 呼んだ: true });
assert.equal(台帳.問い合わせの記録.rows.length, 1 + 5, '台帳の中の問い合わせの記録に見出し＋5行');
assert.equal(台帳.回ごと.rows.length, 1, '回ごとには1回分の1行だけ');
const 行 = 台帳.回ごと.rows[0];
assert.deepEqual(行.slice(1, 4), ['問い合わせの振り分け', '合格', 'AI なし']);
assert.ok(行[5].includes('5 件') && 行[5].includes('Claude を呼んだ'));
assert.equal(台帳.仕事ごと.rows.length, 0, '仕事ごと（式のタブ）には書かない');
const 本物 = global.UrlFetchApp.fetch; global.UrlFetchApp.fetch = () => ({ getResponseCode: () => 500, getContentText: () => 'err' });
台帳.回ごと.rows.length = 0; g.振り分け(); global.UrlFetchApp.fetch = 本物;
assert.equal(台帳.回ごと.rows[0][2], '落ちた', 'つないでいるのに Claude を呼べなかったら「落ちた」');
const url = props.FIRE_URL; delete props.FIRE_URL; delete props.FIRE_TOKEN; 台帳.回ごと.rows.length = 0; g.振り分け(); props.FIRE_URL = url; props.FIRE_TOKEN = 'tok';
assert.equal(台帳.回ごと.rows[0][2], '合格', 'Claude とつないでいない時は振り分けだけで合格');
// FIRE の有無（値は記録しない）。0件でも状態が変わった時だけ1行・半分だけは「落ちた」
const 値を含まない = r => !r.join(' ').includes('tok') && !r.join(' ').includes('trig_x');
inbox.length = 0; 台帳.回ごと.rows.length = 0;
g.振り分け(); 台帳.回ごと.rows.length = 0;   // 直前のテストで変えた状態をここで落ち着かせる
g.振り分け(); assert.equal(台帳.回ごと.rows.length, 0, '状態が同じで0件なら書かない');
delete props.FIRE_TOKEN; g.振り分け();
let r = 台帳.回ごと.rows.pop(); assert.equal(r[2], '落ちた'); assert.ok(r[5].includes('FIRE_TOKEN 未設定') && 値を含まない(r), '半分だけ入っている時は落ちた・値は書かない');
props.FIRE_TOKEN = '  '; g.振り分け(); assert.equal(台帳.回ごと.rows.length, 0, '空白だけは未設定と同じで、状態は変わらない');
delete props.FIRE_URL; delete props.FIRE_TOKEN; g.振り分け();
r = 台帳.回ごと.rows.pop(); assert.equal(r[2], '合格'); assert.ok(r[5].includes('FIRE 未設定'), '両方無いのは「つないでいない」で合格');
props.FIRE_URL = url; props.FIRE_TOKEN = 'tok'; g.振り分け();
r = 台帳.回ごと.rows.pop(); assert.ok(r[5].includes('FIRE 設定済み') && 値を含まない(r), '入れた後の最初の実行で設定済みが1行（値は書かない）');
assert.equal(g.FIREの状態(), 'FIRE 設定済み');
// 「回ごと」が無い表（試しの表）では、最初の1行の時に作る
const 古い = { シート1: タブ('シート1') }; 古い.シート1.rows.push(['見出し']); 開く = 表をつくる(古い);
delete props.FIRE_TOKEN; g.振り分け(); props.FIRE_TOKEN = 'tok';
assert.ok(古い.回ごと && 古い.回ごと.rows[0][0] === '日時' && 古い.回ごと.rows[1][5].includes('FIRE_TOKEN 未設定'), '回ごとを見出しつきで作って書く');
開く = 表をつくる(台帳);
console.log('OK FIRE の有無の記録（値なし）');
// 台帳つきの「Claude とつなぐ」の保存（形だけ確かめる・値は返さない）
global.FIREの状態 = g.FIREの状態; global.会社の設定を当てる = g.会社の設定を当てる;
props.setProperties = null;
global.PropertiesService.getScriptProperties = () => ({ getProperty: k => props[k] || null, setProperty: (k, v) => { props[k] = v; }, setProperties: o => Object.assign(props, o), deleteProperty: k => { delete props[k]; } });
const m = require('./台帳つき/メニュー.js');
delete props.FIRE_URL; delete props.FIRE_TOKEN;
let k = m.鍵を保存('https://example.com/fire', 'x'.repeat(30)); assert.ok(!k.ok && !props.FIRE_URL, '形の違う URL は保存しない');
k = m.鍵を保存(' https://api.anthropic.com/v1/claude_code/routines/trig_abc123/fire ', 'short'); assert.ok(!k.ok && !props.FIRE_TOKEN, '短いトークンは保存しない');
const T = 'sk-ant-' + 'a'.repeat(40);
k = m.鍵を保存('https://api.anthropic.com/v1/claude_code/routines/trig_abc123/fire', ' ' + T + ' ');
assert.ok(k.ok && props.FIRE_TOKEN === T && k.文.includes('FIRE 設定済み') && !k.文.includes(T) && !k.文.includes('trig_abc123'), '保存して有無だけ返す');
// URL を入れてある時（既定 or 保存済み）は、トークンだけで保存できる
k = m.鍵を保存('', T); assert.ok(k.ok && props.FIRE_URL === 'https://api.anthropic.com/v1/claude_code/routines/trig_abc123/fire', 'URL が空なら保存済みの URL を使う');
props.FIRE_URL = url; props.FIRE_TOKEN = 'tok';
console.log('OK Claude とつなぐ（形の確かめ・値を返さない）');
console.log('OK 管理台帳への1行（回ごと・問い合わせの記録）');
fired.length = 0; inbox.length = 0; assert.deepEqual(g.振り分け(), { 全部: 0, AI: 0, 呼んだ: false }); assert.equal(fired.length, 0, '問い合わせが無い時は呼ばない（AI 0）');
// 下書きが最後にあるスレッドは、下書きを除いた最後の1通（問い合わせ）を読む
const 下書き = { getDate: () => new Date(), getSubject: () => '【試し・下書き】Re: 【試し】家賃について', getPlainBody: () => '内見のご案内', isDraft: () => true };
const 問い合わせ = mail('【試し】家賃について', '家賃を教えてください'); 問い合わせ.isDraft = () => false;
受信 = () => [{ getMessages: () => [問い合わせ, 下書き], getMessageCount: () => 2, addLabel: () => {} }];
fired.length = 0; delete props.SHEET_ID; 開く = null;
const nn = g.振り分け(); assert.equal(nn.全部, 1);
assert.ok(fired.length === 1 && JSON.parse(fired[0].o.payload).text.includes('件名：【試し】家賃について') && !JSON.parse(fired[0].o.payload).text.includes('下書き'), '下書きではなく問い合わせを読む');
// 自分が送った物・【試し・下書き】は読まない
global.Session = { getEffectiveUser: () => ({ getEmail: () => 'owner@example.com' }) };
const 自分の = mail('【試し】家賃について', '自分の返信'); 自分の.isDraft = () => false; 自分の.getFrom = () => 'Owner <OWNER@example.com>';
問い合わせ.getFrom = () => 'お客様 <guest@example.net>';
受信 = () => [{ getMessages: () => [問い合わせ, 自分の, 下書き], getMessageCount: () => 3, addLabel: () => {} },
  { getMessages: () => [自分の], getMessageCount: () => 1, addLabel: () => {} }];
fired.length = 0; const n3 = g.振り分け();
assert.equal(n3.全部, 1, '自分だけのスレッドは読まない');
assert.ok(JSON.parse(fired[0].o.payload).text.includes('家賃を教えてください') && !JSON.parse(fired[0].o.payload).text.includes('自分の返信'), 'お客様の1通を読む');
assert.ok(g.設定.探す.includes('-subject:【試し・下書き】'));
問い合わせ.getId = () => 'msg123';
fired.length = 0; g.振り分け(); assert.ok(JSON.parse(fired[0].o.payload).text.includes('メッセージID：msg123'), '元のメッセージの ID を渡す');
console.log('OK 自分が送った物・試しの下書きを読まない');
console.log('OK 下書きを問い合わせとして読まない');
console.log('OK 振り分けのテスト（5通 → AI が要る 3・人へ 2・送信なし）');

// 段1 物件の表：台帳の「物件」のタブを Claude に渡す。無ければ「無し」
{
  const 物件 = タブ('物件'); 物件.getDataRange = () => ({ getValues: () => [['物件', '部屋', '空き', '家賃', '共益費', 'よくある質問'], ['例：みどり東町', '201', '空き', 65000, 8000, '共益費は家賃と別'], ['', '', '', '', '', '']] });
  const 表 = 表をつくる({ 回ごと: タブ('回ごと'), 物件 });
  const t = g.物件の表(表);
  assert.ok(t.includes('答えの根拠はこれだけ') && t.includes('例：みどり東町｜201｜空き｜65000｜8000｜共益費は家賃と別') && !t.includes('｜｜｜'), '物件の表を1行ずつ渡す（空の行は出さない）');
  assert.ok(g.物件の表(表をつくる({ 回ごと: タブ('回ごと') })).includes('無し'), 'タブが無ければ無しと渡す');
  // 振り分けの時に Claude へ渡す文に入る
  props.FIRE_URL = url; props.FIRE_TOKEN = 'tok'; props.SHEET_ID = 'ledger'; 開く = 表;
  受信 = () => [{ getId: () => 'th1', getMessages: () => [問い合わせ], getMessageCount: () => 1, addLabel: () => {} }];
  fired.length = 0; g.振り分け();
  assert.ok(JSON.parse(fired[0].o.payload).text.includes('例：みどり東町'), '振り分けで物件の表を渡す');
  受信 = () => [{ getId: () => 'th1', getMessages: () => [問い合わせ], getMessageCount: () => 2, addLabel: () => {}, getLabels: () => [{ getName: () => 'orbit-済み' }] }];
  fired.length = 0; assert.equal(g.振り分け().全部, 0, '済みの印があるスレッドは、後から下書きが足されて検索に出ても拾い直さない'); assert.equal(fired.length, 0);
  console.log('OK 物件の表を Claude に渡す');
}
// 段3 出口：下書きの控えを取り、送った物と見比べて1行（記録だけ）
{
  const 回 = タブ('回ごと'), 表 = 表をつくる({ 回ごと: 回 }); 開く = 表;
  const 自分 = 'Owner <owner@example.com>';
  const 下 = (txt) => ({ isDraft: () => true, getPlainBody: () => txt, getFrom: () => 自分 });
  const 送 = (txt) => ({ isDraft: () => false, getPlainBody: () => txt, getFrom: () => 自分 });
  const 客 = { isDraft: () => false, getPlainBody: () => '内見したいです', getFrom: () => 'guest@example.net' };
  const スレッド = (id, ms, 古い) => { const t = { id, ms, labels: [], getId: () => id, getMessages: () => t.ms, addLabel: l => t.labels.push(l.name), getLastMessageDate: () => new Date(Date.now() - (古い ? 5 : 0) * 864e5) }; return t; };
  const 本文 = 'A 様\nお問い合わせありがとうございます。\nご都合のよい日時を3つほどお知らせいただけますでしょうか。';
  const a = スレッド('a', [客, 下(本文 + '\n\n2026年10月27日 22:06 山田 が書きました:\n> 内見したいです')]);
  const b = スレッド('b', [客, 下(本文)]);
  const c = スレッド('c', [客, 下(本文)], true);
  出口のスレッド = [a, b, c];
  assert.deepEqual(g.出口を見る(表), [], '下書きを初めて見た回は控えを取るだけ');
  assert.ok(props['draft:a'] === 本文 && !props['draft:a'].includes('書きました'), '控えは引用より前だけ');
  a.ms = [客, 送(本文 + '\n\n2026年10月27日 22:06 山田 が書きました:\n> 内見したいです')];   // そのまま送った
  b.ms = [客, 送(本文.replace('ありがとうございます', 'ありがとうございます。担当の佐藤です'))];   // 直して送った
  c.ms = [客];   // 下書きを消して3日以上たった
  assert.deepEqual(g.出口を見る(表), ['✅そのまま', '✅直して', '❌']);
  assert.deepEqual(回.rows.map(r => r[1] + ' ' + r[2]), ['問い合わせの返信 ✅そのまま', '問い合わせの返信 ✅直して', '問い合わせの返信 ❌']);
  assert.ok(a.labels.includes('orbit-判定済み') && !props['draft:a'] && !props['draft:b'], '判定したら印を付け、控えを消す');
  const d = スレッド('d', [客, 送(本文)]); 出口のスレッド = [d];
  assert.deepEqual(g.出口を見る(表), ['✅（見比べられず）'], '控えを取る前に送った時も1行');
  const 自分の問い = { isDraft: () => false, getPlainBody: () => '内見したいです', getFrom: () => 自分 };
  const e = スレッド('e', [自分の問い, 下(本文)]); 出口のスレッド = [e];
  assert.deepEqual(g.出口を見る(表), [], '問い合わせを自分が送った時（試し）、その1通を返信と見ない');
  e.ms = [自分の問い, 送(本文)];
  assert.deepEqual(g.出口を見る(表), ['✅そのまま'], '返信を送った後に判定する');
  const f = スレッド('f', [自分の問い, 送(本文)]); f.getLabels = () => [{ getName: () => 'orbit-判定済み' }]; 出口のスレッド = [f];
  assert.deepEqual(g.出口を見る(表), [], '判定済みの印があるスレッドは、検索に出ても見ない');
  出口のスレッド = [];
  console.log('OK 出口（そのまま・直して・送らなかったを記録）');
}
// 10/28 読む量：前の回より後の物だけ探す・1日の上限の手前で休んで1回だけ知らせる・落ちたらエラーの文を1日1回残す
{
  const 回 = タブ('回ごと'), 表 = 表をつくる({ 回ごと: 回 }); 開く = 表;
  props.LAST_RUN = String(Date.UTC(2026, 9, 28, 0, 0, 0));
  assert.ok(g.探す条件().endsWith(' after:' + Math.floor((Number(props.LAST_RUN) - 120000) / 1000)), '前の回の2分前より後だけ探す');
  delete props.LAST_RUN; assert.ok(!g.探す条件().includes('after:'), '初回は after を付けない');
  const 今日 = g.今日(); props['READS:' + 今日] = '3000'; 受信 = () => { throw new Error('読んではいけない'); };
  assert.ok(g.振り分け().休み && g.振り分け().休み, '上限に届いたら Gmail を読まずに休む');
  assert.equal(回.rows.filter(r => String(r[5]).includes('上限')).length, 1, '上限の知らせは1日1回');
  delete props['READS:' + 今日]; 受信 = () => { throw new Error('Service invoked too many times for one day: gmail.'); };
  assert.throws(() => g.振り分け()); assert.throws(() => g.振り分け());
  assert.equal(回.rows.filter(r => String(r[5]).includes('too many times') && String(r[5]).includes('分け合って')).length, 1, '落ちた時はエラーの文を1日1回だけ残す');
  受信 = () => []; g.振り分け(); assert.ok(g.読んだ(0) >= 1, '読んだ回数を数える');
  const 止め = タブ('試し：止める'); 止め.getRange = () => ({ getValue: () => 'はい' });
  const 表2 = 表をつくる({ 回ごと: 回, '試し：止める': 止め }); 開く = 表2; props.LAST_RUN = '123'; 受信 = () => { throw new Error('読んではいけない'); };
  assert.ok(g.振り分け().止め && props.LAST_RUN === '123', '「試し：止める」が はい の間は読まず、LAST_RUN も進めない（段4の試し）');
  止め.getRange = () => ({ getValue: () => '' }); 受信 = () => []; g.振り分け(); assert.notEqual(props.LAST_RUN, '123', '空に戻すと動き出す');
  開く = 表;
  let 書いた = null; const 情報 = Object.assign(タブ('Orbit の情報'), { getRange: () => ({ setValues: v => { 書いた = v; } }) });
  global.ScriptApp = Object.assign(global.ScriptApp || {}, { getScriptId: () => 'script123' });
  g.情報を書く(表をつくる({ 'Orbit の情報': 情報 }));
  assert.equal(書いた[0][1], 'script123', '「Orbit の情報」にスクリプトの ID を書く'); assert.ok(/^2026-/.test(書いた[1][1]), '版も書く'); assert.equal(書いた[3][1], 'なし（芯の既定のまま）', 'custom.js が無い時は「なし」');
  書いた = null; delete props.INFO_版; 開く = 表をつくる({ 回ごと: 回, 'Orbit の情報': 情報 }); 受信 = () => [];
  g.振り分け(); assert.ok(書いた && props.INFO_版, '版が変わった後の最初の振り分けで「Orbit の情報」を書き直す');
  書いた = null; g.振り分け(); assert.equal(書いた, null, '同じ版の間は書かない'); 開く = 表;
  console.log('OK 読む量（前の回から・1日の上限・エラーの文）');
}

// 段2 カレンダー：空きから日をずらして3つ・✅ の後に送った文に残る候補だけ仮押さえ
{
  const 予定 = [], 作った = [];
  global.CalendarApp = { getDefaultCalendar: () => ({ getEvents: (a, b) => 予定.filter(e => e.s < b && e.e > a).map(e => ({ getStartTime: () => e.s, getEndTime: () => e.e })), createEvent: (t, s0, s1, o) => 作った.push({ t, s0, s1 }) }) };
  const 今 = new Date(2026, 9, 27, 22, 0);   // 10/27（火）の夜
  予定.push({ s: new Date(2026, 9, 28, 10, 0), e: new Date(2026, 9, 28, 12, 0) });   // 10/28（水）10〜12時は埋まり
  const k = g.空きの候補(今.getTime());
  assert.deepEqual(k.map(x => x.文), ['10月28日（水）12:00〜13:00', '10月29日（木）10:00〜11:00', '10月30日（金）10:00〜11:00'], '埋まりを避けて1日1つ・3つ');
  assert.equal(作った.length, 0, '候補を出すだけではカレンダーに書かない');
  assert.equal(g.仮押さえ('A 様\n' + k[0].文 + '\n' + k[2].文 + '\nのいずれかでいかがでしょうか', k), 2, '送った文に残った2つだけ');
  assert.ok(作った.every(e => e.t === '【自動・仮】内見の候補'));
  console.log('OK カレンダーの候補と、✅ の後だけの仮押さえ');
}
// 段4 止まった時の知らせ：最後に動いてから60分を超えたら1回だけ
{
  const 送った = []; global.MailApp = { sendEmail: (to, sub) => 送った.push({ to, sub }) };
  開く = 表をつくる({ 回ごと: タブ('回ごと') }); props.SHEET_ID = 'ledger';
  props.LAST_RUN = String(Date.now() - 30 * 60000); assert.equal(g.生きてるか見る(), '動いている');
  props.LAST_RUN = String(Date.now() - 90 * 60000); assert.equal(g.生きてるか見る(), '知らせた');
  assert.equal(g.生きてるか見る(), '知らせ済み', '2回目は知らせない');
  assert.ok(送った.length === 1 && 送った[0].to === 'owner@example.com');
  props.LAST_RUN = String(Date.now()); assert.equal(g.生きてるか見る(), '動いている'); assert.ok(!props.STOP_NOTIFIED, '動き出したら印を消す');
  console.log('OK 止まった時の知らせ（1回だけ）');
}

{
  // 会社用の部分（custom.js の 会社の設定）を芯の設定に重ねる。芯の既定はそのまま残り、書いた所だけ変わる（10/30 H の③）
  delete require.cache[require.resolve('./Code.js')];
  global.会社の設定 = { 探す: 'label:問い合わせ', 内見: { 始め: 11, 曜日: [1, 2] }, 既定のFIRE_URL: 'https://example/fire' };
  const c = require('./Code.js'), s = c.会社の設定を当てる();
  assert.ok(s.探す === 'label:問い合わせ' && s.内見.始め === 11 && s.内見.終わり === 18 && s.内見.曜日.join() === '1,2', '書いた所だけ重ねる・配列は置き換え');
  assert.ok(s.ラベル.済み === 'orbit-済み' && s.一日の読む上限 === 3000 && s.既定のFIRE_URL === 'https://example/fire', '書かない所は芯の既定');
  c.会社の設定を当てる(); assert.equal(s.内見.始め, 11, '2回目は重ねない');
  delete global.会社の設定; delete require.cache[require.resolve('./Code.js')];
  assert.equal(require('./Code.js').会社の設定を当てる().探す, g.設定.探す, 'custom.js が無ければ芯の既定のまま');
  console.log('OK 会社用の設定（custom.js）を重ねる');
}

// AI の回数の見積もり（1日・問い合わせ λ 件・Claude の予定の仕事は1時間ごと）
const 空振りの手 = 3, 一件の手 = 6;   // 10/25 の試し：1件の下書きで6手。空振りは「時刻・探す・報告」の3手と見積もり
for (const λ of [3, 5, 10]) {
  const 今 = 24 * 空振りの手 + λ * (一件の手 - 空振りの手);          // 毎時起きる（読む・振り分けも AI）
  const ai件 = Math.round(λ * 0.6);                                  // 振り分けで AI が要るのは返信待ちだけ（見本の割合 3/5）
  const 卒業_毎時 = 24 * 空振りの手 + ai件 * (一件の手 - 空振りの手); // Apps Script が振り分け・AI は毎時起きて返信待ちだけ
  const 卒業_3回 = 3 * 空振りの手 + ai件 * (一件の手 - 空振りの手);  // AI は1日3回だけ起きる（知らせは Apps Script がすぐ出す）
  console.log(`λ=${λ}件/日：今 ${今}手・卒業（AI 毎時）${卒業_毎時}手・卒業（AI 1日3回）${卒業_3回}手（${Math.round(100 - 卒業_3回 / 今 * 100)}%減）`);
}
