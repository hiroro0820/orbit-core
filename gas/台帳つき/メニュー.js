/**
 * 管理台帳の中に入れた形のメニュー（2026-10-27 押す回数を減らす案1〜3・Orbit の窓）
 * 台帳を「コピーを作成」するとこのスクリプトも一緒に付いてくる。SHEET_ID を入れる手は要らない。
 *   Orbit → はじめる      … 許可と、15分ごとの起動と、1回目の振り分けを1回で（案3）
 *   Orbit → Claude とつなぐ … 貼る欄2つの画面で FIRE_URL・FIRE_TOKEN を保存（案2）。名前を打たない。値は記録しない・画面にも戻さない
 *   Orbit → つなぎの状態    … 有無だけを表示
 */
// お客様の Routines の Fire URL（鍵ではない・トークンが無ければ呼べない）。導入の時に私たちが Routines を作ったら、ここに書いてから置く。
// 書いてあれば「Claude とつなぐ」はトークンの欄だけになり、Routines の画面との往復が1回で済む（10/27 案内の順の直し）
var 既定のFIRE_URL = '';

function onOpen() {
  SpreadsheetApp.getUi().createMenu('Orbit')
    .addItem('はじめる（許可と15分ごとの起動）', 'はじめる')
    .addItem('Claude とつなぐ', 'Claudeとつなぐ')
    .addItem('つなぎの状態', 'つなぎの状態')
    .addToUi();
}

function はじめる() {
  起動を入れる();
  情報を書く(表());
  var n = 振り分け();
  SpreadsheetApp.getUi().alert('始めました。これから15分ごとに問い合わせを振り分けます（今回 ' + n.全部 + ' 件）。\nClaude に返信の下書きを作らせる時は、Orbit → Claude とつなぐ を押してください。');
}

function Claudeとつなぐ() {
  var t = HtmlService.createTemplateFromFile('つなぐ');
  t.URLあり = !!(既定のFIRE_URL || 会社の設定を当てる().既定のFIRE_URL || (PropertiesService.getScriptProperties().getProperty('FIRE_URL') || '').trim());
  var html = t.evaluate().setWidth(460).setHeight(t.URLあり ? 260 : 330);
  SpreadsheetApp.getUi().showModalDialog(html, 'Claude とつなぐ');
}

/** 画面から呼ぶ。形だけ確かめて保存し、有無だけを返す（値は返さない・記録しない） */
function 鍵を保存(url, token) {
  url = String(url || '').trim() || 既定のFIRE_URL || 会社の設定を当てる().既定のFIRE_URL || (PropertiesService.getScriptProperties().getProperty('FIRE_URL') || '').trim();
  token = String(token || '').trim();
  if (!/^https:\/\/api\.anthropic\.com\/v1\/claude_code\/routines\/[A-Za-z0-9_]+\/fire$/.test(url))
    return { ok: false, 文: 'URL の形が違います。Routines の画面の「Fire URL」をそのまま貼ってください。' };
  if (token.length < 20 || /\s/.test(token)) return { ok: false, 文: 'トークンの形が違います。Generate token で出た文字をそのまま貼ってください。' };
  PropertiesService.getScriptProperties().setProperties({ FIRE_URL: url, FIRE_TOKEN: token });
  return { ok: true, 文: '保存しました（' + FIREの状態() + '）。この画面は閉じてかまいません。' };
}

function つなぎの状態() {
  SpreadsheetApp.getUi().alert('つなぎの状態：' + FIREの状態());
}

if (typeof module !== 'undefined') module.exports = { onOpen: onOpen, はじめる: はじめる, 鍵を保存: 鍵を保存 };
