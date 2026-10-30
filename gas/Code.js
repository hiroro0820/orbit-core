/**
 * 問い合わせの振り分け（Apps Script の見本・AI なし・送信なし）2026-10-25 Orbit の窓
 * 任せ方の5段目「AI なしにする」の試し：読む・振り分け・ラベル・記録を Google の中で済ませ、
 * AI（Claude の予定の仕事）は「返信待ち」のラベルの物の返信の文だけを書く。
 *   振り分け()   … 新しい問い合わせを読んで振り分ける（15分ごとの時間の起動で動かす）
 *   起動を入れる() … 15分ごとの時間の起動を1つだけ入れる（最初に1回）
 * 返信待ちが1件以上ある時だけ、Claude Code の予定の仕事（Routines）の API の受け口（/fire）を1回呼び、
 * 問い合わせの中身を text で渡す（無い時は AI 0）。受け口の URL とトークンはスクリプトのプロパティ FIRE_URL・FIRE_TOKEN（コードに書かない）。
 * 送らない：メールの送信・転送はしない。Slack への1行は SLACK_WEBHOOK を入れた時だけ（試しでは入れない）。
 */
var ORBIT_版 = '2026-10-30.2';   // 置き直すたびに日付を上げる。台帳の「Orbit の情報」に書き、どのお客様がどの版かを表から見る

var 設定 = {
  探す: 'subject:【試し】 -subject:【試し・下書き】 newer_than:3d -label:orbit-済み',
  自分が送った物を除く: true,   // お客様の問い合わせは自分からは来ない。自分宛ての試しのメールを読ませる時だけ false（custom.js で）
  ラベル: { 返信待ち: 'orbit-返信待ち', 変更: 'orbit-変更・取消', 済み: 'orbit-済み', 判定済み: 'orbit-判定済み' },
  シートの名前: 'orbit 試し：問い合わせの記録',
  SLACK_WEBHOOK: '',   // 試しでは空（送らない）
  // 内見の候補（10/27 段2）：カレンダーの空きから、日をずらして3つ。押さえるのは人が返信を送った後（✅ の後）だけ
  内見: { カレンダー: '', 曜日: [2, 3, 4, 5, 6, 0], 始め: 10, 終わり: 18, 長さ分: 60, 先の日数: 14, 候補の数: 3 },
  止まったら知らせる分: 60,   // 段4：最後に動いてからこれより長いと、自分宛てに1回知らせる
  // Gmail の読み取りは、アカウントの1日の上限（無料のアカウントで約2万回）を、お客様のほかの GAS や連携と分け合う（10/28 同じアカウントの GAS 2つで使い切った）。
  // Orbit が使ってよいのはこの回数まで。届いたら、その日はもう読まずに1回だけ知らせる
  一日の読む上限: 3000,
  一回に開くスレッド: 20
};

/** 会社用の部分（10/30 H の③）：お客様の作業場の custom.js に var 会社の設定 = { … } を置くと、設定に重ねる（物は中まで重ね、配列と値は置き換え）。
 * 芯（このファイル）の更新では custom.js を上書きしない。GAS はファイルの読み込みの順が決まらないので、読み込み時でなく入口の関数の最初に1回だけ重ねる */
var 会社の設定を当てた = false;
function 会社の設定を当てる() {
  if (会社の設定を当てた) return 設定;
  会社の設定を当てた = true;
  var c = typeof 会社の設定 !== 'undefined' ? 会社の設定 : null;
  (function 重ねる(先, 元) {
    Object.keys(元 || {}).forEach(function (k) {
      var v = 元[k];
      if (v && typeof v === 'object' && !Array.isArray(v) && 先[k] && typeof 先[k] === 'object' && !Array.isArray(先[k])) 重ねる(先[k], v);
      else 先[k] = v;
    });
  })(設定, c);
  return 設定;
}

/** 本文と件名から種類を決める（AI なし）。決めきれない物は「返信待ち」に寄せる（AI が文を書く） */
function 種類を決める(件名, 本文) {
  var t = (件名 + ' ' + 本文);
  if (/キャンセル|取り消|取消|やめ/.test(t)) return '取消';
  if (/変更|別の日|日にちを変え|日程を変え/.test(t)) return '変更';
  if (/内見|見学/.test(t)) return '内見';
  return '質問';
}

/** 1通を振り分けた結果（Gmail に触らない部分・テストで使う） */
function 一通を振り分ける(m) {
  var 種類 = 種類を決める(m.件名, m.本文);
  var ラベル = (種類 === '取消' || 種類 === '変更') ? 設定.ラベル.変更 : 設定.ラベル.返信待ち;
  return { 日時: m.日時, 件名: m.件名, 種類: 種類, ラベル: ラベル, AIが要る: ラベル === 設定.ラベル.返信待ち };
}

function ラベル(name) {
  return GmailApp.getUserLabelByName(name) || GmailApp.createLabel(name);
}

/** 記録のシート。SHEET_ID が管理台帳（「回ごと」のタブがある表）なら、その中の「問い合わせの記録」のタブ（無ければ作る）。
 * 無ければ前と同じく新しい表を作る */
function 表() {
  // 管理台帳の中に入れた形（台帳つき/）なら、その台帳そのもの。SHEET_ID は要らない（10/27 押す回数を減らす案1）
  var 台帳 = SpreadsheetApp.getActiveSpreadsheet ? SpreadsheetApp.getActiveSpreadsheet() : null;
  if (台帳) return 台帳;
  var p = PropertiesService.getScriptProperties(), id = p.getProperty('SHEET_ID'), ss;
  if (id) { try { ss = SpreadsheetApp.openById(id); } catch (e) { ss = null; } }
  if (!ss) { ss = SpreadsheetApp.create(設定.シートの名前); p.setProperty('SHEET_ID', ss.getId()); }
  return ss;
}

function シート(ss) {
  var 見出し = ['記録した日時', '受けた日時', '件名', '種類', '付けたラベル', 'AI が要るか'], sh;
  if (ss.getSheetByName && ss.getSheetByName('回ごと')) {
    sh = ss.getSheetByName('問い合わせの記録');
    if (!sh) { sh = ss.insertSheet('問い合わせの記録'); sh.appendRow(見出し); }
    return sh;
  }
  sh = ss.getSheets()[0];
  if (sh.getLastRow() === 0) sh.appendRow(見出し);
  return sh;
}

/** 管理台帳の「回ごと」に1回分の1行（日時・仕事・結果・任せ方の段・かかった秒・メモ）。
 * 「回ごと」のタブが無い表（管理台帳につなぐ前の試しの表）には、見出しつきで作ってから書く（つなぎの有無を AI が値を見ずに確かめられるように・10/27） */
function 台帳に1行(ss, 結果, 秒, メモ) {
  if (!ss.getSheetByName) return;
  var sh = ss.getSheetByName('回ごと');
  if (!sh) { sh = ss.insertSheet('回ごと'); sh.appendRow(['日時', '仕事', '結果', '任せ方の段', 'かかった秒', 'メモ']); }
  sh.appendRow([new Date(), '問い合わせの振り分け', 結果, 'AI なし', 秒, メモ]);
}

/** 台帳の「物件」のタブを、Claude に渡す文にする（10/27 段1）。Claude はこれだけを根拠に答え、無い事は「担当者が確認」。タブが無ければ「無し」と渡す */
function 物件の表(ss) {
  var sh = ss.getSheetByName && ss.getSheetByName('物件');
  if (!sh) return '物件の表：無し（物件・空き・料金はすべて「担当者が確認のうえご連絡いたします」とする）';
  var rows = sh.getDataRange().getValues().filter(function (r) { return r.join('').trim(); });
  var 文 = rows.map(function (r) { return r.map(function (c) { return String(c).replace(/\s+/g, ' ').trim(); }).join('｜'); }).join('\n');
  return '物件の表（答えの根拠はこれだけ。1行目が見出し）：\n' + 文.slice(0, 4000);
}

/** 出口（10/27 段3・記録だけ）：返信待ちのスレッドで、人が下書きを「直さず送った／直して送った／送らなかった」を見て、回ごとに1行。人は何も書かない。
 * 下書きを初めて見た時に本文の控えを取り、送った物が出たら見比べる。3日たっても送らず下書きも消えていれば「送らなかった」 */
function 出口を見る(ss) {
  会社の設定を当てる();
  var p = PropertiesService.getScriptProperties(), 判定 = [];
  // ふだんは前の回より後に動いた（下書きができた・送った）スレッドだけ。3日送らなかった（❌）を見るため、1日1回だけ4日分を見る。14日・50件を毎回開いていた（10/28 上限の件）
  var 一日一回 = p.getProperty('EXIT_FULL') !== 今日(); if (一日一回) p.setProperty('EXIT_FULL', 今日());
  読んだ(1);
  GmailApp.search('label:' + 設定.ラベル.返信待ち + ' -label:' + 設定.ラベル.判定済み + ' newer_than:4d' + (一日一回 ? '' : 前の回から()), 0, 設定.一回に開くスレッド).forEach(function (th) {
    読んだ(1);
    if (印がある(th, 設定.ラベル.判定済み)) return;   // 振り分けと同じ理由で、検索だけでは判定済みを外せない
    var id = th.getId(), ms = th.getMessages(), 控え = p.getProperty('draft:' + id);
    読んだ(ms.length);
    var 下書き = ms.filter(function (m) { return m.isDraft && m.isDraft(); }).pop();
    if (下書き && !控え) { p.setProperty('draft:' + id, 本文だけ(下書き.getPlainBody()).slice(0, 3000)); p.setProperty('from:' + id, 差出人(下書き.getFrom())); 控え = 'x'; }
    // 自分＝下書きの差出人（自分のアドレスを読む許可を足さずに済む）。控えが無い時だけ自分のアドレスを読む（読めなければ空）
    var 自分 = p.getProperty('from:' + id) || 自分のアドレス();
    // 1通目は問い合わせなので除く（試しは問い合わせも自分が送るので、除かないと返信前に「送った」と判定した・10/27 23:21 本物の試し）
    var 送った = ms.filter(function (m, i) { return i > 0 && !(m.isDraft && m.isDraft()) && 自分 && 差出人(m.getFrom()) === 自分; }).pop();
    if (!p.getProperty('draft:' + id) && !送った) return;   // まだ下書きができていない
    var 結果 = null;
    if (送った && !p.getProperty('draft:' + id)) 結果 = '✅（見比べられず）';   // 控えを取る前に送った（15分の見回りより早く送った時）。続けて直しなしの数には入れない
    else if (送った) 結果 = 本文だけ(送った.getPlainBody()) === p.getProperty('draft:' + id) ? '✅そのまま' : '✅直して';
    else if (!下書き && (Date.now() - th.getLastMessageDate().getTime()) > 3 * 24 * 3600 * 1000) 結果 = '❌';
    if (!結果) return;
    var 押さえ = 0, slots = p.getProperty('slots:' + id);
    if (送った && slots) { try { 押さえ = 仮押さえ(送った.getPlainBody(), JSON.parse(slots)); } catch (e) { 押さえ = -1; } }
    判定.push(結果); th.addLabel(ラベル(設定.ラベル.判定済み)); p.deleteProperty('draft:' + id); p.deleteProperty('from:' + id); p.deleteProperty('slots:' + id);
    var sh = ss.getSheetByName && ss.getSheetByName('回ごと');
    if (sh) sh.appendRow([new Date(), '問い合わせの返信', 結果, '半分', '', (結果 === '✅そのまま' ? '下書きを直さず送った' : 結果 === '✅直して' ? '下書きを直して送った' : 結果 === '❌' ? '3日送らず下書きも無い' : '下書きの控えを取る前に送った') + (押さえ > 0 ? '・内見の候補を ' + 押さえ + ' 枠 仮押さえ' : 押さえ < 0 ? '・仮押さえに失敗' : '')]);
  });
  return 判定;
}

/** スレッドにその印（ラベル）があるか。テストの作り物で getLabels が無い時は「無い」 */
function 印がある(th, 名前) { return !!(th.getLabels && th.getLabels().some(function (l) { return l.getName() === 名前; })); }

/** 「名前 <a@b>」から a@b だけ（小文字） */
function 差出人(f) { var m = String(f || '').match(/<([^>]+)>/); return (m ? m[1] : String(f || '')).trim().toLowerCase(); }

/** 自分のアドレス。読む許可（userinfo.email）が無い時は空を返し、止まらない */
function 自分のアドレス() {
  try { return typeof Session !== 'undefined' ? String(Session.getEffectiveUser().getEmail() || '').toLowerCase() : ''; } catch (e) { return ''; }
}

var 曜 = ['日', '月', '火', '水', '木', '金', '土'];
function 日時の文(d, 分) {
  var e = new Date(d.getTime() + 分 * 60000), hm = function (x) { return x.getHours() + ':' + ('0' + x.getMinutes()).slice(-2); };
  return (d.getMonth() + 1) + '月' + d.getDate() + '日（' + 曜[d.getDay()] + '）' + hm(d) + '〜' + hm(e);
}

function 内見のカレンダー() {
  return 設定.内見.カレンダー ? CalendarApp.getCalendarById(設定.内見.カレンダー) : CalendarApp.getDefaultCalendar();
}

/** 段2：内見の候補。明日から 先の日数 の間で、内見の曜日・時間帯の中の、予定と重ならない枠を、1日1つずつ 候補の数 まで。カレンダーには書かない */
function 空きの候補(今) {
  var c = 設定.内見, cal = 内見のカレンダー(), out = [], 今日 = new Date(今 || Date.now());
  for (var i = 1; i <= c.先の日数 && out.length < c.候補の数; i++) {
    var 日 = new Date(今日.getFullYear(), 今日.getMonth(), 今日.getDate() + i);
    if (c.曜日.indexOf(日.getDay()) < 0) continue;
    var 予定 = cal.getEvents(new Date(日.getTime() + c.始め * 3600000), new Date(日.getTime() + c.終わり * 3600000));
    for (var h = c.始め; h + c.長さ分 / 60 <= c.終わり; h++) {
      var s0 = new Date(日.getTime() + h * 3600000), s1 = new Date(s0.getTime() + c.長さ分 * 60000);
      var 重なる = 予定.some(function (e) { return e.getStartTime() < s1 && e.getEndTime() > s0; });
      if (!重なる) { out.push({ 始め: s0.toISOString(), 文: 日時の文(s0, c.長さ分) }); break; }
    }
  }
  return out;
}

/** 段2：人が返信を送った後（✅）、送った文に残っている候補だけを【自動・仮】で押さえる。人が消した候補は押さえない */
function 仮押さえ(送った文, 候補) {
  var cal = 内見のカレンダー(), n = 0;
  (候補 || []).forEach(function (k) {
    if (String(送った文).indexOf(k.文) < 0) return;
    var s0 = new Date(k.始め);
    cal.createEvent('【自動・仮】内見の候補', s0, new Date(s0.getTime() + 設定.内見.長さ分 * 60000),
      { description: 'Orbit が、返信で伺った候補を押さえました。お客様の返事が来たら1つに決めて、残りは消してください。' });
    n++;
  });
  return n;
}

/** 段4：止まった時の知らせ（1時間ごとの起動で動く）。最後に動いてから 止まったら知らせる分 より長ければ、自分宛てに1回だけ知らせ、回ごとに1行。動き出したら印を消す */
function 生きてるか見る() {
  会社の設定を当てる();
  var p = PropertiesService.getScriptProperties(), last = Number(p.getProperty('LAST_RUN') || 0), 分 = (Date.now() - last) / 60000;
  if (!last) return 'まだ1回も動いていない';
  if (分 <= 設定.止まったら知らせる分) { p.deleteProperty('STOP_NOTIFIED'); return '動いている'; }
  if (p.getProperty('STOP_NOTIFIED')) return '知らせ済み';
  var 自分 = 自分のアドレス(), 文 = '問い合わせの振り分けが ' + Math.round(分) + ' 分動いていません。台帳を開いて Orbit → はじめる を押すと、もう一度動き出します。';
  if (自分) MailApp.sendEmail(自分, '【Orbit】問い合わせの振り分けが止まっています', 文);
  台帳に1行(表(), '落ちた', '', '止まっている：' + Math.round(分) + ' 分動いていない' + (自分 ? '（メールで知らせた）' : '（知らせ先が分からない）'));
  p.setProperty('STOP_NOTIFIED', String(Date.now()));
  return '知らせた';
}

/** 返信の本文だけ（引用より前・空白をそろえる）。見比べに使う */
function 本文だけ(t) {
  var lines = String(t || '').replace(/\r/g, '').split('\n'), out = [];
  for (var i = 0; i < lines.length; i++) {
    var l = lines[i];
    if (/^>/.test(l) || /(wrote:|書きました[:：]?)\s*$/.test(l)) break;
    out.push(l.trim());
  }
  return out.join('\n').replace(/\n{2,}/g, '\n').trim();
}

/** 今日の Gmail の読み取りの回数（スクリプトのプロパティに日ごと）。n を足して、足した後の数を返す */
function 読んだ(n) {
  var p = PropertiesService.getScriptProperties(), k = 'READS:' + 今日(), v = Number(p.getProperty(k) || 0) + (n || 0);
  if (n) p.setProperty(k, String(v));
  return v;
}
function 今日() { var d = new Date(); return d.getFullYear() + '-' + (d.getMonth() + 1) + '-' + d.getDate(); }

/** 1日1回だけ表に1行・メールで知らせる（同じ種類の知らせを15分ごとに重ねない） */
function 今日1回だけ知らせる(ss, 種類, 文) {
  var p = PropertiesService.getScriptProperties(), k = 'NOTICE:' + 種類 + ':' + 今日();
  if (p.getProperty(k)) return false;
  p.setProperty(k, '1');
  台帳に1行(ss, '落ちた', '', 文);
  var 自分 = 自分のアドレス();
  try { if (自分) MailApp.sendEmail(自分, '【Orbit】' + 文.slice(0, 40), 文); } catch (e) {}
  return true;
}

/** 15分ごとの起動で呼ばれる。上限の見張りとエラーの記録をして、中身（振り分けの本体）を呼ぶ */
/** 段4の本物の試し用：「試し：止める」タブの A1 が「はい」の間は、何もせず LAST_RUN も進めない（＝止まった時と同じ）。タブが無い台帳（お客様）には効かない。メニューはスマホで出ないので、セル1つで切り替える（10/28） */
function 試しに止めている(ss) {
  var sh = ss.getSheetByName && ss.getSheetByName('試し：止める');
  return !!sh && String(sh.getRange('A1').getValue()).trim() === 'はい';
}

function 振り分け() {
  会社の設定を当てる();
  var ss = 表(), p = PropertiesService.getScriptProperties();
  if (試しに止めている(ss)) return { 全部: 0, AI: 0, 呼んだ: false, 止め: true };
  if (読んだ(0) >= 設定.一日の読む上限) {
    今日1回だけ知らせる(ss, '上限', '今日の Gmail の読み取りが Orbit の上限（' + 設定.一日の読む上限 + ' 回）に届いたので、明日まで振り分けを休みます。ほかの仕組みの分を残すためです。');
    p.setProperty('LAST_RUN', String(Date.now()));   // 自分で休んでいるだけなので、止まった知らせ（段4）は出さない
    return { 全部: 0, AI: 0, 呼んだ: false, 休み: true };
  }
  if (p.getProperty('INFO_版') !== ORBIT_版) { try { 情報を書く(ss); p.setProperty('INFO_版', ORBIT_版); } catch (e) {} }   // 新しい版を当てた後の最初の回で「Orbit の情報」を書き直す（作業場が表で版を確かめる・10/30 H）
  try { return 振り分けの本体(ss); }
  catch (e) {
    今日1回だけ知らせる(ss, 'エラー', '振り分けが落ちました：' + String(e && e.message || e).slice(0, 200) + (/too many times/i.test(String(e && e.message)) ? '（Gmail の1日の上限。お客様のほかの GAS や連携と分け合っています）' : ''));
    throw e;
  }
}

/** 前の回より後にメッセージが来たスレッドだけ読む（after:<前の回の秒−2分>・projects/oubo-gas と同じ形・10/28）。初回は付けない */
function 前の回から() {
  var last = Number(PropertiesService.getScriptProperties().getProperty('LAST_RUN') || 0);
  return last ? ' after:' + Math.floor((last - 120000) / 1000) : '';
}
function 探す条件() { return 設定.探す + 前の回から(); }

function 振り分けの本体(ss) {
  var 始め = Date.now(), threads = GmailApp.search(探す条件(), 0, 設定.一回に開くスレッド), sh = シート(ss), n = { 全部: 0, AI: 0, 呼んだ: false }, 待ち = [], 候補 = null;
  threads.forEach(function (th) {
    読んだ(1);
    if (印がある(th, 設定.ラベル.済み)) return;   // 検索は1通ずつ見るので、後から足された下書き（印なし）があると済みのスレッドも出てくる（10/27 23:51 本物の試しで同じスレッドを15分ごとに拾い直した）
    // スレッドの最後の1通が下書き（Claude が作った返信の下書きなど）のことがあるので、下書きを除いた最後の1通を読む（10/27 本物の試しで、下書きを問い合わせとして拾った）
    var 自分 = 設定.自分が送った物を除く ? 自分のアドレス() : '';
    var ms = th.getMessages().filter(function (m) {
      if (m.isDraft && m.isDraft()) return false;
      if (/^【試し・下書き】/.test(m.getSubject())) return false;
      if (自分 && m.getFrom && 差出人(m.getFrom()) === 自分) return false;
      return true;
    }), msg = ms[ms.length - 1];
    読んだ(ms.length + 1);
    if (!msg) return;
    var r = 一通を振り分ける({ 日時: msg.getDate(), 件名: msg.getSubject(), 本文: msg.getPlainBody().slice(0, 2000) });
    th.addLabel(ラベル(r.ラベル)); th.addLabel(ラベル(設定.ラベル.済み));
    sh.appendRow([new Date(), r.日時, r.件名, r.種類, r.ラベル, r.AIが要る ? 'はい' : 'いいえ（人へ）']);
    // メッセージID は、元のスレッドへの返信の下書きにするため（10/27）
    var 候補文 = '';
    if (r.AIが要る && r.種類 === '内見') {
      try { 候補 = 候補 || 空きの候補(); } catch (e) { 候補 = []; }
      if (候補.length) { 候補文 = '\n内見の候補（カレンダーの空き。返信ではこの文のまま並べる）：' + 候補.map(function (k) { return k.文; }).join('／'); if (th.getId) PropertiesService.getScriptProperties().setProperty('slots:' + th.getId(), JSON.stringify(候補)); }
    }
    n.全部++; if (r.AIが要る) { n.AI++; 待ち.push('メッセージID：' + (msg.getId ? msg.getId() : '') + '\n件名：' + r.件名 + 候補文 + '\n本文：' + msg.getPlainBody().slice(0, 1500)); }
  });
  if (待ち.length) n.呼んだ = Claudeを呼ぶ('問い合わせ ' + 待ち.length + '件（Apps Script の振り分けで「返信待ち」）\n' + 物件の表(ss) + '\n---\n' + 待ち.join('\n---\n'));
  出口を見る(ss);
  if (n.全部 && 設定.SLACK_WEBHOOK) {
    UrlFetchApp.fetch(設定.SLACK_WEBHOOK, { method: 'post', contentType: 'application/json',
      payload: JSON.stringify({ text: '問い合わせ ' + n.全部 + ' 件（返信の下書きが要る ' + n.AI + ' 件）' }) });
  }
  // つなぎの状態（有無だけ）。半分だけ入っている時は入れ間違いなので「落ちた」。状態が前の回から変わった時は、0件でも1行残す（最初の実行で確かめられるように）
  var 状態 = FIREの状態(), p = PropertiesService.getScriptProperties(), 変わった = p.getProperty('LAST_FIRE_STATE') !== 状態;
  var 半分 = 状態 === 'FIRE_URL 未設定' || 状態 === 'FIRE_TOKEN 未設定', 落ちた = 半分 || (n.AI && 状態 === 'FIRE 設定済み' && !n.呼んだ);
  if (n.全部 || 変わった) 台帳に1行(ss, 落ちた ? '落ちた' : '合格', Math.round((Date.now() - 始め) / 1000),
    n.全部 + ' 件・返信の下書きが要る ' + n.AI + ' 件・' + 状態 + (!n.AI ? '' : n.呼んだ ? '（Claude を呼んだ）' : 落ちた ? '（Claude を呼べなかった）' : '（Claude とつないでいない）'));
  p.setProperty('LAST_FIRE_STATE', 状態);
  p.setProperty('LAST_RUN', String(Date.now()));   // 段4の見張りが読む
  console.log('振り分け：' + n.全部 + ' 件・AI が要る ' + n.AI + ' 件');
  return n;
}

/** FIRE_URL・FIRE_TOKEN の有無だけを言葉で返す（値は返さない・記録しない）。空白だけの値は無いと同じ */
function FIREの状態() {
  var p = PropertiesService.getScriptProperties(), u = (p.getProperty('FIRE_URL') || '').trim(), t = (p.getProperty('FIRE_TOKEN') || '').trim();
  return !u && !t ? 'FIRE 未設定' : !u ? 'FIRE_URL 未設定' : !t ? 'FIRE_TOKEN 未設定' : 'FIRE 設定済み';
}

/** Routines の /fire を1回呼ぶ。プロパティが無ければ呼ばない（false）。公式：code.claude.com/docs/en/routines */
function Claudeを呼ぶ(text) {
  var p = PropertiesService.getScriptProperties(), url = (p.getProperty('FIRE_URL') || '').trim(), token = (p.getProperty('FIRE_TOKEN') || '').trim();
  if (!url || !token) { console.log('FIRE_URL・FIRE_TOKEN が無いので Claude は呼ばない'); return false; }
  var res = UrlFetchApp.fetch(url, { method: 'post', contentType: 'application/json', muteHttpExceptions: true,
    headers: { Authorization: 'Bearer ' + token, 'anthropic-beta': 'experimental-cc-routine-2026-04-01', 'anthropic-version': '2023-06-01' },
    payload: JSON.stringify({ text: text.slice(0, 20000) }) });
  console.log('Claude を呼んだ：' + res.getResponseCode() + ' ' + res.getContentText().slice(0, 300));
  return res.getResponseCode() === 200;
}

/** 「Orbit の情報」タブに、スクリプトの ID・版・はじめた日時を書く（はじめるの時）。後から直した版を置き直す時に、ID を探す手を要らなくする（10/27 試しの台帳で ID が分からずオーナーの手が要った） */
function 情報を書く(ss) {
  if (!ss.getSheetByName) return;
  var sh = ss.getSheetByName('Orbit の情報') || ss.insertSheet('Orbit の情報');
  会社の設定を当てる();
  var 会社 = typeof 会社の設定 !== 'undefined' && 会社の設定 ? Object.keys(会社の設定).length : 0;
  sh.getRange(1, 1, 4, 2).setValues([['スクリプトの ID', ScriptApp.getScriptId()], ['版', ORBIT_版], ['はじめた日時', new Date()],
    ['会社用の部分', 会社 ? 'あり（' + 会社 + ' 項目・custom.js）' : 'なし（芯の既定のまま）']]);
}

function 起動を入れる() {
  会社の設定を当てる();
  ScriptApp.getProjectTriggers().forEach(function (t) { if (t.getHandlerFunction() === '振り分け') ScriptApp.deleteTrigger(t); });
  ScriptApp.newTrigger('振り分け').timeBased().everyMinutes(15).create();
  ScriptApp.getProjectTriggers().forEach(function (t) { if (t.getHandlerFunction() === '生きてるか見る') ScriptApp.deleteTrigger(t); });
  ScriptApp.newTrigger('生きてるか見る').timeBased().everyHours(1).create();
}

if (typeof module !== 'undefined') module.exports = { 種類を決める: 種類を決める, 一通を振り分ける: 一通を振り分ける, 振り分け: 振り分け, 設定: 設定, 会社の設定を当てる: 会社の設定を当てる, FIREの状態: FIREの状態, 物件の表: 物件の表, 出口を見る: 出口を見る, 本文だけ: 本文だけ, 空きの候補: 空きの候補, 仮押さえ: 仮押さえ, 生きてるか見る: 生きてるか見る, 日時の文: 日時の文, 情報を書く: 情報を書く, 試しに止めている: 試しに止めている, 探す条件: 探す条件, 今日: 今日, 読んだ: 読んだ };
