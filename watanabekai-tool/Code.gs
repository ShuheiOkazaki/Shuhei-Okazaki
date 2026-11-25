/**
 * 渡辺ゼミ 卒業論文発表会・渡辺会 運営ツール
 *
 * 使い方: 空のGoogleスプレッドシートを開き、拡張機能 > Apps Script に
 * このファイルを貼り付けて保存 → シートを再読み込み →
 * メニュー「渡辺会ツール」>「① 初期セットアップ」を実行。
 *
 * 作られるシート
 *   設定            開催日などの入力欄（毎年ここだけ更新）
 *   発表者          発表者・論文タイトルを入力（発表時間は自動計算）
 *   式次第          設定と発表者から自動生成される式次第（印刷/PDF化用）
 *   TODO            準備タスク（期限は開催日から自動計算）
 *   出欠フォーム設問 出欠フォームの設問定義（編集可）
 *   審査フォーム設問 発表者別フォームの設問定義（編集可）
 */

const SHEET = {
  SET: '設定', PRES: '発表者', PROG: '式次第', TODO: 'TODO',
  QA: '出欠フォーム設問', QR: '審査フォーム設問',
};
const SET_FIRST = 4;   // 設定シートの先頭データ行
const PRES_FIRST = 5;  // 発表者シートの先頭データ行
const PRES_MAX = 12;   // 発表者の最大枠
const PRES_LAST = PRES_FIRST + PRES_MAX - 1;
const TODO_FIRST = 6;
const TODO_LAST = 60;
const Q_FIRST = 6;     // 設問シートの先頭設問行

/* ================= メニュー ================= */

function onOpen() {
  SpreadsheetApp.getUi().createMenu('渡辺会ツール')
    .addItem('① 初期セットアップ（シート作成）', 'setupAll')
    .addSeparator()
    .addItem('② 出欠フォームを作成', 'createAttendanceForm')
    .addItem('③ 発表者別の審査フォームを一括作成', 'createReviewForms')
    .addItem('   選択行の審査フォームを作り直す', 'recreateSelectedReviewForms')
    .addItem('⑤ 審査結果を集計（点数・順位）', 'updateSummary')
    .addItem('⑥ 不要な回答タブを整理（削除）', 'cleanupResponseSheets')
    .addSeparator()
    .addItem('④ 新年度の準備（入力のリセット）', 'resetForNewYear')
    .addToUi();
}

/* ================= 設定の定義 ================= */

/**
 * t: date | time | num | text | long | choice
 * role: in(入力) | calc(自動計算) | out(自動記入)
 * v: 初期値（"=..." は数式）。関数なら S(キー) でほかの設定セルを参照できる
 */
function settingsDefs_() {
  return [
    { sec: '■ 基本（毎年ここを更新）' },
    { k: '開催日', t: 'date', role: 'in', v: '=DATE(2027,1,31)', n: '★毎年ここを変更。TODOの期限なども自動で連動' },
    { k: '期', t: 'num', role: 'in', v: 22, n: '今年度卒業する期（発表者の期）' },
    { k: '年度', t: 'num', role: 'calc', v: S => `=IF(MONTH(${S('開催日')})<4,YEAR(${S('開催日')})-1,YEAR(${S('開催日')}))`, n: '4月始まりで自動計算' },
    { k: '曜日', t: 'text', role: 'calc', v: S => `=CHOOSE(WEEKDAY(${S('開催日')}),"日","月","火","水","木","金","土")` },
    { k: '日付表示', t: 'text', role: 'calc', v: S => `=YEAR(${S('開催日')})&"年"&MONTH(${S('開催日')})&"月"&DAY(${S('開催日')})&"日（"&${S('曜日')}&"）"` },
    { k: '現役生の集合対象', t: 'text', role: 'calc', v: S => `=${S('期')}&"期・"&(${S('期')}+1)&"期・"&(${S('期')}+2)&"期"` },

    { sec: '■ 卒業論文発表会' },
    { k: '会場名', t: 'text', role: 'in', v: '中央大学 茗荷谷キャンパス' },
    { k: '会場URL', t: 'text', role: 'in', v: 'https://www.chuo-u.ac.jp/access/myogadani/' },
    { k: '教室', t: 'text', role: 'in', v: '地下1階 B1C16 教室' },
    { k: '現役生集合時刻', t: 'time', role: 'in', v: '=TIME(12,30,0)' },
    { k: '開会時刻', t: 'time', role: 'in', v: '=TIME(14,15,0)', n: '開会の辞の開始時刻' },
    { k: '開会の辞(分)', t: 'num', role: 'in', v: 10 },
    { k: '報告時間(分)', t: 'num', role: 'in', v: 15 },
    { k: '質疑時間(分)', t: 'num', role: 'in', v: 10 },
    { k: '発表間隔(分)', t: 'num', role: 'in', v: 5, n: '発表と発表の入替時間' },
    { k: '休憩を入れる発表番号', t: 'num', role: 'in', v: 3, n: 'この番号の発表の後に休憩（不要なら 0）' },
    { k: '休憩(分)', t: 'num', role: 'in', v: 15 },
    { k: '卒論発表会終了', t: 'time', role: 'calc', v: S => `=IF(COUNT('${SHEET.PRES}'!$E$${PRES_FIRST}:$E$${PRES_LAST})=0,"",MAX('${SHEET.PRES}'!$E$${PRES_FIRST}:$E$${PRES_LAST}))`, n: '最後の発表の終了時刻（発表者シートから自動）' },
    { k: '資料配布', t: 'text', role: 'in', v: 'LINEグループで配布します。' },
    { k: '問合せ先', t: 'text', role: 'in', v: '【担当者のメールアドレス・電話番号を入力】' },

    { sec: '■ 渡辺会・懇親会' },
    { k: '総会開始', t: 'time', role: 'in', v: '=TIME(18,0,0)' },
    { k: '総会終了', t: 'time', role: 'in', v: '=TIME(19,0,0)' },
    { k: '総会内容', t: 'text', role: 'in', v: '開会挨拶，会の趣旨説明，各期の紹介等' },
    { k: '懇親会場所', t: 'text', role: 'in', v: '茗荷谷キャンパス内スエヒロ' },
    { k: '懇親会開始', t: 'time', role: 'in', v: '=TIME(19,0,0)' },
    { k: '懇親会終了', t: 'time', role: 'in', v: '', n: '任意。※昨年は見積書が18:00〜20:00表記だったため、お店に要確認' },
    { k: 'OB・OG会費(円)', t: 'num', role: 'in', v: 6000, n: '暫定の場合は設問文側で「（暫定）」と補足' },
    { k: '人数変更期限', t: 'date', role: 'in', v: S => `=${S('開催日')}-16`, n: '初期値は開催日の16日前。上書き可' },

    { sec: '■ 受講報告会（ぴあスポーツビジネススクール）' },
    { k: '受講報告会を実施', t: 'choice', role: 'in', v: 'はい', n: 'はい／いいえ。「いいえ」だと式次第・出欠フォームから外れる' },
    { k: '受講場所', t: 'text', role: 'in', v: S => `=${S('教室')}` },
    { k: '受講開始', t: 'time', role: 'in', v: '=TIME(13,0,0)' },
    { k: '受講終了', t: 'time', role: 'in', v: '=TIME(14,0,0)' },
    { k: '受講報告者', t: 'text', role: 'in', v: '渡辺ゼミ○期生 ○○○○' },
    { k: '受講趣旨', t: 'long', role: 'in', v: '渡辺ゼミOBOGあるいは現役生の中からの選抜者は，毎年1，2名，きらぼし銀行様のご厚意により受講料を支弁していただき，通常は20数万円かかる「ぴあビジネススクール」を無料で受講させていただいています。例年，卒業論文発表会の際に，数か月に渡る受講内容につき，受講者から報告がなされることとなっています。この報告会には，きらぼし銀行様ならびにぴあからもご出席いただきます。企業の皆様のご厚志に報いるためにも，渡辺ゼミ現役生は全員この報告会に参加しなければなりません。もちろん，OBOG各位の中で関心がある人は，参加可能です。' },

    { sec: '■ フォーム設定' },
    { k: '審査回答をこのシートに集約', t: 'choice', role: 'in', v: 'はい', n: 'はい: 各審査フォームの回答がこのスプレッドシートに「審査_○」シートで溜まる' },

    { sec: '■ 自動記入（触らない）' },
    { k: '出欠フォームURL', t: 'text', role: 'out', v: '', n: '回答者に配るURL' },
    { k: '出欠フォーム編集URL', t: 'text', role: 'out', v: '' },
    { k: 'フォーム保存フォルダURL', t: 'text', role: 'out', v: '' },
  ];
}

function makeRef_() {
  const rowOf = {};
  settingsDefs_().forEach((d, i) => { if (d.k) rowOf[d.k] = SET_FIRST + i; });
  const S = k => {
    if (!rowOf[k]) throw new Error('設定キーが見つかりません: ' + k);
    return `'${SHEET.SET}'!$B$${rowOf[k]}`;
  };
  S.row = k => rowOf[k];
  return S;
}

/* ================= セットアップ ================= */

function setupAll() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActive();
  const exists = Object.values(SHEET).filter(n => ss.getSheetByName(n));
  if (exists.length) {
    const r = ui.alert('既存シートの上書き確認',
      '次のシートは内容が初期化されます（入力済みデータは消えます）:\n' + exists.join(', ') + '\n\n続けますか？',
      ui.ButtonSet.YES_NO);
    if (r !== ui.Button.YES) return;
  }
  const shs = {};
  Object.keys(SHEET).forEach(key => {
    let sh = ss.getSheetByName(SHEET[key]);
    if (!sh) sh = ss.insertSheet(SHEET[key]);
    sh.clear();
    sh.setConditionalFormatRules([]);
    sh.getRange(1, 1, sh.getMaxRows(), sh.getMaxColumns()).breakApart();
    shs[key] = sh;
  });
  const S = makeRef_();
  buildSettings_(shs.SET, S);
  buildPresenters_(shs.PRES, S);
  buildProgram_(shs.PROG, S);
  buildTodo_(shs.TODO, S);
  buildAttendanceQ_(shs.QA);
  buildReviewQ_(shs.QR);
  // 並び順
  ['SET', 'PRES', 'PROG', 'TODO', 'QA', 'QR'].forEach((key, i) => {
    ss.setActiveSheet(shs[key]);
    ss.moveActiveSheet(i + 1);
  });
  const dflt = ss.getSheetByName('シート1') || ss.getSheetByName('Sheet1');
  if (dflt && ss.getSheets().length > 1 && dflt.getLastRow() === 0) ss.deleteSheet(dflt);
  ss.setActiveSheet(shs.SET);
  ui.alert('セットアップ完了。「設定」シートの黄色いセルを入力してください。');
}

const C_IN = '#fff8dc', C_CALC = '#eeeeee', C_HEAD = '#1f4e79', C_SEC = '#d9e2f3';

function styleHeader_(range) {
  range.setBackground(C_HEAD).setFontColor('#ffffff').setFontWeight('bold')
    .setHorizontalAlignment('center').setVerticalAlignment('middle');
  return range;
}

/* ---------- 設定 ---------- */
function buildSettings_(sh, S) {
  const defs = settingsDefs_();
  sh.getRange('A1').setValue('設定（毎年、黄色いセルを更新）').setFontSize(14).setFontWeight('bold');
  sh.getRange('A2').setValue('黄色=入力欄 / 灰色=自動計算（編集不要）。最初に開催日を入れると、TODOの期限・出欠フォームの日付が連動します。')
    .setFontColor('#666666');
  defs.forEach((d, i) => {
    const row = SET_FIRST + i;
    if (d.sec) {
      sh.getRange(row, 1, 1, 3).setBackground(C_SEC).setFontWeight('bold');
      sh.getRange(row, 1).setValue(d.sec);
      return;
    }
    sh.getRange(row, 1).setValue(d.k).setFontWeight('bold');
    const cell = sh.getRange(row, 2);
    const v = typeof d.v === 'function' ? d.v(S) : d.v;
    if (typeof v === 'string' && v.charAt(0) === '=') cell.setFormula(v);
    else if (v !== '') cell.setValue(v);
    cell.setBackground(d.role === 'in' ? C_IN : C_CALC).setHorizontalAlignment('left');
    if (d.t === 'date') cell.setNumberFormat('yyyy"年"m"月"d"日"');
    if (d.t === 'time') cell.setNumberFormat('H:mm');
    if (d.t === 'num') cell.setNumberFormat('#,##0');
    if (d.t === 'long') cell.setWrap(true);
    if (d.t === 'choice') {
      cell.setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(['はい', 'いいえ'], true).build());
    }
    if (d.n) sh.getRange(row, 3).setValue(d.n).setFontColor('#666666');
  });
  sh.setColumnWidth(1, 220);
  sh.setColumnWidth(2, 420);
  sh.setColumnWidth(3, 520);
  sh.setFrozenRows(2);
}

/* ---------- 発表者 ---------- */
function buildPresenters_(sh, S) {
  sh.getRange('A1').setValue('発表者（氏名・論文タイトルを入力 → 発表時間と式次第が自動更新）').setFontSize(14).setFontWeight('bold');
  sh.getRange('A2').setValue('連名は「松川拓真・加藤夏実」のように「・」でつなぐ。上から詰めて入力。入力後に メニュー「渡辺会ツール」>「③ 審査フォームを一括作成」。')
    .setFontColor('#666666');
  const head = ['番号', '発表者氏名', '論文タイトル', '開始', '終了', '審査フォームURL（配布用）', '審査フォーム編集URL', '状態', '式次第用(名簿)', '式次第用(時程)'];
  styleHeader_(sh.getRange(4, 1, 1, head.length).setValues([head]));
  for (let i = 0; i < PRES_MAX; i++) {
    const r = PRES_FIRST + i;
    sh.getRange(r, 1).setValue(i + 1);
    const D = i === 0
      ? `=IF($B${r}="","",${S('開会時刻')}+${S('開会の辞(分)')}/1440)`
      : `=IF($B${r}="","",E${r - 1}+IF($A${r - 1}=${S('休憩を入れる発表番号')},${S('休憩(分)')},${S('発表間隔(分)')})/1440)`;
    sh.getRange(r, 4).setFormula(D);
    sh.getRange(r, 5).setFormula(`=IF($B${r}="","",D${r}+(${S('報告時間(分)')}+${S('質疑時間(分)')})/1440)`);
    sh.getRange(r, 8).setFormula(`=IF($B${r}="","",IF($F${r}="","未作成","作成済"))`);
    sh.getRange(r, 9).setFormula(`=IF($B${r}="","",UNICHAR(9311+$A${r})&" "&$B${r})`);
    sh.getRange(r, 10).setFormula(
      `=IF($B${r}="","",UNICHAR(9311+$A${r})&" "&TEXT(D${r},"H:mm")&"～"&TEXT(E${r},"H:mm")&" "&$B${r}` +
      `&IF($C${r}="","","「"&$C${r}&"」")` +
      `&IF(AND($A${r}=${S('休憩を入れる発表番号')},$B${r + 1}<>""),CHAR(10)&"休憩 "&TEXT(E${r},"H:mm")&"～"&TEXT(E${r}+${S('休憩(分)')}/1440,"H:mm"),""))`);
  }
  sh.getRange(PRES_FIRST, 2, PRES_MAX, 2).setBackground(C_IN);
  sh.getRange(PRES_FIRST, 4, PRES_MAX, 2).setNumberFormat('H:mm').setBackground(C_CALC);
  sh.getRange(PRES_FIRST, 6, PRES_MAX, 3).setBackground(C_CALC);
  sh.getRange(PRES_FIRST, 9, PRES_MAX, 2).setBackground(C_CALC).setFontColor('#999999');
  sh.getRange(PRES_FIRST, 1, PRES_MAX, 1).setHorizontalAlignment('center');
  sh.getRange(PRES_FIRST, 1, PRES_MAX, 10).setVerticalAlignment('middle');
  [60, 220, 300, 70, 70, 330, 330, 80, 160, 160].forEach((w, i) => sh.setColumnWidth(i + 1, w));
  sh.setFrozenRows(4);
}

/* ---------- 式次第 ---------- */
function buildProgram_(sh, S) {
  const P = c => `'${SHEET.PRES}'!${c}`;
  const T = k => `TEXT(${S(k)},"H:mm")`;
  const lines = [
    { f: `=${S('年度')}&"年度 渡辺ゼミ "&${S('期')}&"期生 卒業論文発表会"`, style: 'title' },
    { f: `="■ 日時："&${S('日付表示')}&${T('開会時刻')}&"～"&${T('卒論発表会終了')}` },
    { f: `="■ 場所："&${S('会場名')}&"（"&${S('会場URL')}&"）・"&${S('教室')}` },
    { f: `="■ 現役生の集合時間："&${T('現役生集合時刻')}&"（遅刻厳禁！！）"&${S('現役生の集合対象')}&"は全員この時間に集合！！"`, style: 'warn' },
    { f: `="■ 資料："&${S('資料配布')}` },
    { f: `="■ 報告時間："&${S('報告時間(分)')}&"分　質疑応答時間："&${S('質疑時間(分)')}&"分"` },
    { f: '■ 審査シート：後日グーグルフォームのアドレスを記載したヴァージョンをお送りします！' },
    { f: `=TEXTJOIN(CHAR(10),TRUE,${P(`$I$${PRES_FIRST}:$I$${PRES_LAST}`)})`, style: 'indent' },
    { f: '■ 発表者と発表時間', style: 'bold' },
    { f: `="（開会の辞）"&${T('開会時刻')}&"～"&TEXT(${S('開会時刻')}+${S('開会の辞(分)')}/1440,"H:mm")` },
    { f: `=TEXTJOIN(CHAR(10),TRUE,${P(`$J$${PRES_FIRST}:$J$${PRES_LAST}`)})` },
    { f: '' },
    { f: '渡辺会と懇親会', style: 'center' },
    { f: `="■ 渡辺会総会："&${T('総会開始')}&"～"&${T('総会終了')}&"　"&${S('総会内容')}`, style: 'warn' },
    { f: '■ 卒業論文発表会の最優秀賞の発表：以下の懇親会内' },
    { f: `="■ 懇親会（"&${S('懇親会場所')}&"）："&${T('懇親会開始')}&"～"&IF(${S('懇親会終了')}="","",${T('懇親会終了')})&"　参加者は現役生とOBOGを想定"`, style: 'warn' },
    { f: `="※ 問合せ先 "&${S('問合せ先')}` },
    { f: '' },
    { f: `=IF(${S('受講報告会を実施')}="はい","ぴあスポーツビジネススクール 受講報告会","")`, style: 'center' },
    { f: `=IF(${S('受講報告会を実施')}="はい","■ 趣旨"&CHAR(10)&${S('受講趣旨')},"")` },
    { f: `=IF(${S('受講報告会を実施')}="はい","■ 場所："&${S('会場名')}&"・"&${S('受講場所')},"")` },
    { f: `=IF(${S('受講報告会を実施')}="はい","■ 時間："&${T('受講開始')}&"～"&${T('受講終了')},"")` },
    { f: `=IF(${S('受講報告会を実施')}="はい","■ 報告者："&${S('受講報告者')},"")` },
  ];
  sh.setColumnWidth(1, 20);
  sh.setColumnWidth(2, 780);
  lines.forEach((ln, i) => {
    const cell = sh.getRange(i + 1, 2);
    if (ln.f.charAt(0) === '=') cell.setFormula(ln.f); else cell.setValue(ln.f);
    cell.setWrap(true).setVerticalAlignment('top').setFontSize(11);
    if (ln.style === 'title') cell.setFontSize(16).setFontWeight('bold').setHorizontalAlignment('center');
    if (ln.style === 'center') cell.setFontSize(13).setFontWeight('bold').setHorizontalAlignment('center');
    if (ln.style === 'bold') cell.setFontWeight('bold');
    if (ln.style === 'warn') cell.setFontColor('#c00000').setFontWeight('bold');
  });
  sh.setHiddenGridlines(true);
}

/* ---------- TODO ---------- */
function buildTodo_(sh, S) {
  sh.getRange('A1').setValue('準備TODO（期限は「開催日」から自動計算）').setFontSize(14).setFontWeight('bold');
  sh.getRange('B2').setFormula(`="進捗： "&COUNTIF(G${TODO_FIRST}:G${TODO_LAST},"完了")&" / "&COUNTA(C${TODO_FIRST}:C${TODO_LAST})&" 件完了"`);
  sh.getRange('B3').setFormula(`="開催まであと "&(${S('開催日')}-TODAY())&" 日（開催日："&${S('日付表示')}&"）"`);
  sh.getRange('B4').setFormula(`="期限超過： "&COUNTIFS(F${TODO_FIRST}:F${TODO_LAST},"<"&TODAY(),G${TODO_FIRST}:G${TODO_LAST},"<>完了")&" 件"`);
  sh.getRange('B2:B4').setFontWeight('bold');
  const head = ['No', '区分', 'タスク', '担当', '期限（開催日から○日）', '期限日', 'ステータス', '備考'];
  styleHeader_(sh.getRange(5, 1, 1, head.length).setValues([head]));

  const A = '① 出欠確認・参加者募集', B = '② 当日運営準備', C = '③ 当日対応', D = '④ 事後';
  const tasks = [
    [A, 'Googleフォーム作成（卒論発表会・渡辺会・懇親会の出欠）', -60, 'メニュー「② 出欠フォームを作成」で自動生成'],
    [A, '回答集計用スプレッドシート作成・共有', -60, '出欠フォームの回答先はこのブック内に自動作成'],
    [A, '懇親会の会場（スエヒロ）の時間・見積書の確認', -60, '昨年は見積書が18:00〜20:00表記。開始・終了時間を要確認（追加提案）'],
    [A, 'OB・OG向け案内文作成', -58, ''],
    [A, '各期の連絡担当へ案内文・フォーム共有', -55, ''],
    [A, '各期LINE等で参加案内を展開', -55, ''],
    [A, '未回答者・参加未定者へのリマインド（1回目）', -35, ''],
    [A, '未回答者・参加未定者へのリマインド（2回目）', -24, ''],
    [A, '出欠集計・渡辺先生への人数報告', -17, ''],
    [A, '人数変更期限までに最終人数を確認・共有', -16, '「設定」の人数変更期限と連動'],
    [B, '発表者の確定・設定シートへの入力', -45, '発表者シートに入力 → 式次第が自動更新'],
    [B, '発表者別の審査フォーム作成・配布', -30, 'メニュー「③ 審査フォームを一括作成」'],
    [B, '渡辺会の司会担当決定', -45, ''],
    [B, '開会挨拶・会の説明担当の確認', -45, ''],
    [B, '総会用スライドの準備・更新', -21, ''],
    [B, '司会進行内容の確認', -14, ''],
    [B, '現役生へ集合時間・式次第を周知', -14, '追加提案'],
    [B, '参加者への最終案内', -7, ''],
    [C, '受付・参加者確認', 0, ''],
    [C, '懇親会費の集金', 0, ''],
    [C, '渡辺会総会の司会進行', 0, ''],
    [D, '懇親会費の精算・会計報告', 7, '追加提案'],
    [D, '次年度への引継ぎ（担当・反省点の記録）', 14, '追加提案'],
  ];
  const rows = tasks.map((t, i) => [i + 1, t[0], t[1], '', t[2], '', '未着手', t[3]]);
  sh.getRange(TODO_FIRST, 1, rows.length, 8).setValues(rows);
  for (let r = TODO_FIRST; r <= TODO_LAST; r++) {
    sh.getRange(r, 6).setFormula(`=IF(E${r}="","",${S('開催日')}+E${r})`);
  }
  sh.getRange(TODO_FIRST, 6, TODO_LAST - TODO_FIRST + 1, 1).setNumberFormat('yyyy/m/d(ddd)');
  sh.getRange(TODO_FIRST, 6, TODO_LAST - TODO_FIRST + 1, 1).setBackground(C_CALC);
  sh.getRange(TODO_FIRST, 7, TODO_LAST - TODO_FIRST + 1, 1).setDataValidation(
    SpreadsheetApp.newDataValidation().requireValueInList(['未着手', '進行中', '完了'], true).build());
  sh.getRange(TODO_FIRST, 4, TODO_LAST - TODO_FIRST + 1, 1).setBackground(C_IN);
  sh.getRange(TODO_FIRST, 1, TODO_LAST - TODO_FIRST + 1, 8).setVerticalAlignment('middle');
  sh.getRange(TODO_FIRST, 3, TODO_LAST - TODO_FIRST + 1, 1).setWrap(true);
  sh.getRange(TODO_FIRST, 8, TODO_LAST - TODO_FIRST + 1, 1).setWrap(true);
  [40, 170, 380, 110, 90, 120, 90, 380].forEach((w, i) => sh.setColumnWidth(i + 1, w));
  sh.setFrozenRows(5);

  const rng = sh.getRange(TODO_FIRST, 1, TODO_LAST - TODO_FIRST + 1, 8);
  const rule = (f, bg, fc, strike) => {
    const b = SpreadsheetApp.newConditionalFormatRule().whenFormulaSatisfied(f).setRanges([rng]);
    if (bg) b.setBackground(bg);
    if (fc) b.setFontColor(fc);
    if (strike) b.setStrikethrough(true);
    return b.build();
  };
  sh.setConditionalFormatRules([
    rule(`=$G${TODO_FIRST}="完了"`, '#e7e6e6', '#999999', true),
    rule(`=AND($F${TODO_FIRST}<>"",$F${TODO_FIRST}<TODAY(),$G${TODO_FIRST}<>"完了")`, '#f8cbad', '#c00000', false),
    rule(`=AND($F${TODO_FIRST}<>"",$F${TODO_FIRST}-TODAY()<=7,$F${TODO_FIRST}>=TODAY(),$G${TODO_FIRST}<>"完了")`, '#fff2cc', null, false),
    rule(`=$G${TODO_FIRST}="進行中"`, '#ddebf7', null, false),
  ]);
}

/* ---------- 設問シート ---------- */
function buildQSheet_(sh, title, desc, confirm, questions) {
  sh.getRange('A1:A3').setValues([['フォームタイトル'], ['フォーム説明'], ['送信後メッセージ']]).setFontWeight('bold');
  sh.getRange('B1').setValue(title);
  sh.getRange('B2').setValue(desc).setWrap(true);
  sh.getRange('B3').setValue(confirm).setWrap(true);
  sh.getRange('B1:G1').merge(); sh.getRange('B2:G2').merge(); sh.getRange('B3:G3').merge();
  sh.getRange('B1:B3').setBackground(C_IN).setVerticalAlignment('top');
  sh.getRange('A4').setValue('{設定のキー名} は設定シートの値に置換されます（例 {日付表示} {期} {会場名}）。{期一覧}=1期〜現役最終期。選択肢はカンマ区切り。目盛は「最小,最大,左ラベル,右ラベル」。')
    .setFontColor('#666666');
  const head = ['順', '設問', '種類', '選択肢 / 目盛', '必須', '説明文', '表示条件', '配点の重み（審査のみ）'];
  styleHeader_(sh.getRange(5, 1, 1, head.length).setValues([head]));
  sh.getRange(Q_FIRST, 1, questions.length, 7).setValues(questions.map((q, i) => [i + 1].concat(q)));
  const n = 40;
  sh.getRange(Q_FIRST, 2, n, 6).setBackground(C_IN).setWrap(true).setVerticalAlignment('top');
  sh.getRange(Q_FIRST, 3, n, 1).setDataValidation(SpreadsheetApp.newDataValidation()
    .requireValueInList(['記述式', '段落', 'ラジオ', 'チェックボックス', 'プルダウン', '目盛', 'セクション'], true).build());
  sh.getRange(Q_FIRST, 5, n, 1).setDataValidation(SpreadsheetApp.newDataValidation()
    .requireValueInList(['○'], true).build()).setHorizontalAlignment('center');
  sh.getRange(Q_FIRST, 7, n, 1).setDataValidation(SpreadsheetApp.newDataValidation()
    .requireValueInList(['受講報告会'], true).setAllowInvalid(false).build());
  [40, 420, 120, 300, 50, 320, 110, 130].forEach((w, i) => sh.setColumnWidth(i + 1, w));
  sh.setFrozenRows(5);
}

function buildAttendanceQ_(sh) {
  const att = '出席,欠席,未定';
  buildQSheet_(sh,
    '{年度}年度 渡辺ゼミ{期}期生 卒業論文発表会・渡辺会・懇親会 出欠確認',
    '{日付表示}に開催する卒業論文発表会・渡辺会総会・懇親会の出欠確認フォームです。\n' +
    '場所：{会場名}（{教室}）\n' +
    '卒業論文発表会：{開会時刻}～{卒論発表会終了}\n' +
    '渡辺会総会：{総会開始}～{総会終了}\n' +
    '懇親会：{懇親会開始}～（{懇親会場所}）\n' +
    '回答・変更の期限：{人数変更期限}',
    'ご回答ありがとうございました。当日お会いできることを楽しみにしています。',
    [
      ['氏名', '記述式', '', '○', '', ''],
      ['期', 'プルダウン', '{期一覧}', '○', '', ''],
      ['区分', 'ラジオ', '現役生,OB・OG', '○', '', ''],
      ['連絡先メールアドレス', '記述式', '', '', '当日の連絡用', ''],
      ['出欠', 'セクション', '', '', '', ''],
      ['卒業論文発表会（{開会時刻}～{卒論発表会終了}）', 'ラジオ', att, '○', '会場：{会場名} {教室}', ''],
      ['ぴあスポーツビジネススクール 受講報告会（{受講開始}～{受講終了}）', 'ラジオ', att, '○', '現役生は全員参加。OB・OGは任意', '受講報告会'],
      ['渡辺会総会（{総会開始}～{総会終了}）', 'ラジオ', att, '○', '', ''],
      ['懇親会（{懇親会開始}～ {懇親会場所}）', 'ラジオ', att, '○', 'OB・OG会費：約{OB・OG会費(円)}円（暫定）', ''],
      ['アレルギー・食事の制限など', '段落', '', '', '', ''],
      ['連絡事項・ご質問', '段落', '', '', '', ''],
    ]);
}

function buildReviewQ_(sh) {
  const sc = '1,5,低い,高い';
  buildQSheet_(sh,
    '{年度}年度 渡辺ゼミ{期}期生 卒業論文発表会 審査シート【{番号} {発表者}】',
    '発表者：{発表者}\n論文タイトル：{論文タイトル}\n発表時間：{発表時間}',
    '審査のご回答ありがとうございました。',
    [
      ['審査者の氏名', '記述式', '', '○', '', ''],
      ['審査者の区分', 'ラジオ', '現役生,OB・OG,教員・来賓', '○', '', ''],
      ['テーマの独創性・意義', '目盛', sc, '○', '', ''],
      ['論理構成・議論の妥当性', '目盛', sc, '○', '', ''],
      ['データ・分析手法の適切さ', '目盛', sc, '○', '', ''],
      ['発表の分かりやすさ（資料・話し方）', '目盛', sc, '○', '', ''],
      ['質疑応答の的確さ', '目盛', sc, '○', '', ''],
      ['総合評価', '目盛', '1,10,低い,高い', '○', '', ''],
      ['良かった点', '段落', '', '', '', ''],
      ['改善点・コメント', '段落', '', '', '', ''],
    ]);
}

/* ================= フォーム生成 ================= */

function readSettings_() {
  const sh = SpreadsheetApp.getActive().getSheetByName(SHEET.SET);
  if (!sh) throw new Error('先に「① 初期セットアップ」を実行してください。');
  const defs = settingsDefs_();
  const vals = sh.getRange(SET_FIRST, 1, defs.length, 2).getDisplayValues();
  const map = {};
  vals.forEach(r => { if (r[0] && r[0].charAt(0) !== '■') map[r[0]] = r[1]; });
  const S = makeRef_();
  const dateVal = sh.getRange(S.row('開催日'), 2).getValue();
  map.__dateOk = dateVal instanceof Date && !isNaN(dateVal);
  const last = parseInt(map['期'], 10);
  const list = [];
  for (let i = 1; i <= last + 2; i++) list.push(i + '期');
  map['期一覧'] = list.join(',');
  return map;
}

function setSetting_(key, value) {
  const sh = SpreadsheetApp.getActive().getSheetByName(SHEET.SET);
  sh.getRange(makeRef_().row(key), 2).setValue(value);
}

function applyPh_(text, map) {
  return String(text == null ? '' : text).replace(/\{([^}]+)\}/g, (m, k) =>
    Object.prototype.hasOwnProperty.call(map, k) ? map[k] : m);
}

function readQuestions_(sheetName) {
  const sh = SpreadsheetApp.getActive().getSheetByName(sheetName);
  if (!sh) throw new Error('シートがありません: ' + sheetName);
  const meta = sh.getRange(1, 2, 3, 1).getValues().map(r => r[0]);
  if (sheetName === SHEET.QR && !sh.getRange(5, 8).getValue()) {
    styleHeader_(sh.getRange(5, 8).setValue('配点の重み（審査のみ）'));
    sh.setColumnWidth(8, 130);
    sh.getRange(Q_FIRST, 8, 40, 1).setBackground(C_IN).setHorizontalAlignment('center');
  }
  const n = Math.max(sh.getLastRow() - Q_FIRST + 1, 0);
  const rows = n ? sh.getRange(Q_FIRST, 1, n, 8).getValues() : [];
  const items = rows.filter(r => String(r[1]).trim() !== '').map(r => ({
    title: String(r[1]), type: String(r[2]).trim(), choices: String(r[3]),
    required: String(r[4]).trim() === '○', help: String(r[5]), cond: String(r[6]).trim(),
    weight: parseFloat(r[7]) > 0 ? parseFloat(r[7]) : 1,
  }));
  return { title: meta[0], desc: meta[1], confirm: meta[2], items: items };
}

function buildForm_(form, q, map) {
  form.setTitle(applyPh_(q.title, map));
  form.setDescription(applyPh_(q.desc, map));
  if (q.confirm) form.setConfirmationMessage(applyPh_(q.confirm, map));
  try { form.setCollectEmail(false); } catch (e) { /* 組織設定で不可の場合は無視 */ }
  try { form.setRequireLogin(false); } catch (e) { /* 同上 */ }
  q.items.forEach(it => {
    if (it.cond === '受講報告会' && map['受講報告会を実施'] !== 'はい') return;
    addItem_(form, it, map);
  });
}

function addItem_(form, it, map) {
  const title = applyPh_(it.title, map);
  const help = applyPh_(it.help, map);
  const choices = applyPh_(it.choices, map).split(/[,、\n]/).map(s => s.trim()).filter(Boolean);
  let item;
  switch (it.type) {
    case '記述式': item = form.addTextItem(); break;
    case '段落': item = form.addParagraphTextItem(); break;
    case 'ラジオ': item = form.addMultipleChoiceItem().setChoiceValues(choices); break;
    case 'チェックボックス': item = form.addCheckboxItem().setChoiceValues(choices); break;
    case 'プルダウン': item = form.addListItem().setChoiceValues(choices); break;
    case '目盛': {
      const lo = parseInt(choices[0] || '1', 10), hi = parseInt(choices[1] || '5', 10);
      item = form.addScaleItem().setBounds(lo, hi);
      if (choices[2] || choices[3]) item.setLabels(choices[2] || '', choices[3] || '');
      break;
    }
    case 'セクション':
      item = form.addSectionHeaderItem().setTitle(title);
      if (help) item.setHelpText(help);
      return;
    default: throw new Error('未対応の設問種類です: 「' + it.type + '」（設問: ' + title + '）');
  }
  item.setTitle(title);
  if (help) item.setHelpText(help);
  if (it.required) item.setRequired(true);
}

function getFolder_(map) {
  const name = '渡辺会_' + map['年度'] + '年度_フォーム';
  const it = DriveApp.getFoldersByName(name);
  const folder = it.hasNext() ? it.next() : DriveApp.createFolder(name);
  setSetting_('フォーム保存フォルダURL', folder.getUrl());
  return folder;
}

function shortUrl_(form) {
  const long = form.getPublishedUrl();
  try { return form.shortenFormUrl(long) || long; } catch (e) { return long; }
}

/** フォームの回答先をこのスプレッドシートにし、できたシートを rename する */
function attachDestination_(form, sheetName) {
  const ss = SpreadsheetApp.getActive();
  const before = ss.getSheets().map(s => s.getSheetId());
  form.setDestination(FormApp.DestinationType.SPREADSHEET, ss.getId());
  SpreadsheetApp.flush();
  Utilities.sleep(1500);
  const added = ss.getSheets().filter(s => before.indexOf(s.getSheetId()) < 0);
  if (added.length) {
    let name = sheetName, i = 2;
    while (ss.getSheetByName(name)) name = sheetName + '_' + i++;
    try { added[0].setName(name); } catch (e) { /* 名前変更失敗は無視 */ }
  }
}

function createAttendanceForm() {
  const ui = SpreadsheetApp.getUi();
  const map = readSettings_();
  if (!map.__dateOk) { ui.alert('「設定」シートの開催日が日付になっていません。'); return; }
  if (map['出欠フォームURL']) {
    const r = ui.alert('作成済みです',
      '出欠フォームは作成済みです。新しく作り直しますか？\n（古いフォームと回答はGoogleドライブに残ります）',
      ui.ButtonSet.YES_NO);
    if (r !== ui.Button.YES) return;
  }
  const form = FormApp.create(applyPh_(readQuestions_(SHEET.QA).title, map));
  buildForm_(form, readQuestions_(SHEET.QA), map);
  DriveApp.getFileById(form.getId()).moveTo(getFolder_(map));
  attachDestination_(form, '出欠回答' + map['年度']);
  setSetting_('出欠フォームURL', shortUrl_(form));
  setSetting_('出欠フォーム編集URL', form.getEditUrl());
  ui.alert('出欠フォームを作成しました。URLは「設定」シートに記入しています。');
}

function createReviewForms() { createReviewForms_(false, null); }

function recreateSelectedReviewForms() {
  const sh = SpreadsheetApp.getActiveSheet();
  if (sh.getName() !== SHEET.PRES) {
    SpreadsheetApp.getUi().alert('「発表者」シートで作り直したい行を選択してから実行してください。');
    return;
  }
  const rng = sh.getActiveRange();
  const rows = [];
  for (let r = rng.getRow(); r < rng.getRow() + rng.getNumRows(); r++) {
    if (r >= PRES_FIRST && r <= PRES_LAST) rows.push(r);
  }
  createReviewForms_(true, rows);
}

function createReviewForms_(force, onlyRows) {
  const ui = SpreadsheetApp.getUi();
  const map = readSettings_();
  const sh = SpreadsheetApp.getActive().getSheetByName(SHEET.PRES);
  const q = readQuestions_(SHEET.QR);
  const data = sh.getRange(PRES_FIRST, 1, PRES_MAX, 7).getDisplayValues();
  const rows = [];
  data.forEach((d, i) => {
    const r = PRES_FIRST + i;
    if (!d[1].trim()) return;
    if (onlyRows && onlyRows.indexOf(r) < 0) return;
    if (d[5] && !force) return;
    rows.push(r);
  });
  if (!rows.length) { ui.alert('作成対象がありません（氏名が空、または作成済み）。'); return; }
  const hasOld = rows.some(r => data[r - PRES_FIRST][6]);
  const ok = ui.alert('審査フォーム作成',
    rows.length + ' 件の審査フォームを作成します。' +
    (hasOld ? '\n\n※作り直す発表者の古い回答タブは削除されます（回答データは古いフォーム側に残ります）。' : '') +
    '\nよろしいですか？', ui.ButtonSet.YES_NO);
  if (ok !== ui.Button.YES) return;
  const folder = getFolder_(map);
  const toSheet = map['審査回答をこのシートに集約'] === 'はい';
  let done = 0;
  rows.forEach(r => {
    const d = data[r - PRES_FIRST];
    const num = String.fromCharCode(9311 + parseInt(d[0], 10));
    const m = Object.assign({}, map, {
      '番号': num, '発表者': d[1].trim(),
      '論文タイトル': d[2] || '（未定）', '発表時間': d[3] && d[4] ? d[3] + '～' + d[4] : '',
    });
    if (d[6]) removeResponseSheetsOf_(formIdOf_(d[6]));
    const form = FormApp.create(applyPh_(q.title, m));
    buildForm_(form, q, m);
    DriveApp.getFileById(form.getId()).moveTo(folder);
    if (toSheet) attachDestination_(form, '審査_' + d[0] + '_' + d[1].trim().slice(0, 20));
    sh.getRange(r, 6).setValue(shortUrl_(form));
    sh.getRange(r, 7).setValue(form.getEditUrl());
    done++;
  });
  ui.alert(done + ' 件の審査フォームを作成しました。「発表者」シートにURLを記入しています。');
}

/* ================= 回答タブの整理 ================= */

function formIdOf_(url) {
  const m = String(url || '').match(/\/forms\/d\/(?!e\/)([^\/?#]+)/);
  return m ? m[1] : '';
}

function linkedFormId_(sheet) {
  let url = '';
  try { url = sheet.getFormUrl(); } catch (e) { return ''; }
  if (!url) return '';
  const id = formIdOf_(url);
  if (id) return id;
  try { return FormApp.openByUrl(url).getId(); } catch (e) { return ''; }
}

/** 指定フォームに紐づく回答タブの紐づけを解除して削除する。削除したタブ名を返す */
function removeResponseSheetsOf_(formId) {
  const ss = SpreadsheetApp.getActive();
  const removed = [];
  if (!formId) return removed;
  ss.getSheets().filter(sh => linkedFormId_(sh) === formId).forEach(sh => {
    const name = sh.getName();
    try { FormApp.openById(formId).removeDestination(); } catch (e) { /* フォーム削除済みなど */ }
    try {
      SpreadsheetApp.flush();
      ss.deleteSheet(sh);
      removed.push(name);
    } catch (e) { /* 削除できなければ手動で「フォームの紐付けを解除」 */ }
  });
  return removed;
}

/** 現在使っているフォーム（出欠・各審査）以外に紐づく回答タブを一覧にして、確認のうえ削除する */
function cleanupResponseSheets() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActive();
  const keep = {};
  const map = readSettings_();
  keep[formIdOf_(map['出欠フォームURL'] ? map['出欠フォーム編集URL'] : '')] = true;
  const pres = ss.getSheetByName(SHEET.PRES);
  pres.getRange(PRES_FIRST, 7, PRES_MAX, 1).getDisplayValues().forEach(r => { keep[formIdOf_(r[0])] = true; });
  const targets = ss.getSheets().filter(sh => {
    const id = linkedFormId_(sh);
    return id && !keep[id];
  });
  if (!targets.length) { ui.alert('削除対象の回答タブはありません。'); return; }
  const r = ui.alert('不要な回答タブの削除',
    '現在のフォームに紐づいていない次の回答タブを削除します。\n\n' +
    targets.map(sh => '・' + sh.getName()).join('\n') +
    '\n\n※フォームとの紐づけを解除して削除します。回答データ自体は各フォーム側に残ります。\n削除しますか？',
    ui.ButtonSet.YES_NO);
  if (r !== ui.Button.YES) return;
  const failed = [];
  targets.forEach(sh => {
    const name = sh.getName();
    const id = linkedFormId_(sh);
    try { FormApp.openById(id).removeDestination(); } catch (e) { /* 無視 */ }
    try { SpreadsheetApp.flush(); ss.deleteSheet(sh); } catch (e) { failed.push(name); }
  });
  ui.alert('削除しました。' + (failed.length ? '\n\n削除できなかったタブ（手動で対応してください）:\n' + failed.join('\n') : ''));
}

/* ================= 新年度リセット ================= */

function resetForNewYear() {
  const ui = SpreadsheetApp.getUi();
  const r = ui.alert('新年度の準備',
    '次を初期化します。\n・発表者の氏名／タイトル／フォームURL\n・出欠フォームURL\n・TODOのステータス（全て「未着手」）\n\n' +
    '※回答シート（出欠回答○○／審査_○）とドライブ上のフォームは残ります。必要なら先に別名で保存してください。\n' +
    '※「設定」シートの開催日・期などは手で更新してください。続けますか？', ui.ButtonSet.YES_NO);
  if (r !== ui.Button.YES) return;
  const ss = SpreadsheetApp.getActive();
  const pres = ss.getSheetByName(SHEET.PRES);
  pres.getRange(PRES_FIRST, 2, PRES_MAX, 2).clearContent();
  pres.getRange(PRES_FIRST, 6, PRES_MAX, 2).clearContent();
  const todo = ss.getSheetByName(SHEET.TODO);
  const last = todo.getRange(TODO_FIRST, 3, TODO_LAST - TODO_FIRST + 1, 1).getValues().filter(v => v[0] !== '').length;
  if (last) todo.getRange(TODO_FIRST, 7, last, 1).setValue('未着手');
  setSetting_('出欠フォームURL', '');
  setSetting_('出欠フォーム編集URL', '');
  setSetting_('フォーム保存フォルダURL', '');
  ui.alert('初期化しました。「設定」シートの開催日・期を更新してください。');
}

/* ================= 審査結果の集計 ================= */

const SUM_SHEET = '審査集計';
const CMT_SHEET = '審査コメント';

/**
 * 発表者ごとの審査フォームの回答を読み、項目別平均・総合得点(100点満点換算)・順位を
 * 「審査集計」シートに、自由記述を「審査コメント」シートに書き出す。何度でも再実行可。
 * 総合得点 = 目盛項目ごとに (平均/最大値) を配点の重みで加重平均 ×100
 */
function updateSummary() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActive();
  const pres = ss.getSheetByName(SHEET.PRES);
  if (!pres) { ui.alert('先に「① 初期セットアップ」を実行してください。'); return; }
  const map = readSettings_();
  const q = readQuestions_(SHEET.QR);
  const scales = q.items.filter(it => it.type === '目盛').map(it => {
    const c = applyPh_(it.choices, map).split(/[,、\n]/).map(x => x.trim());
    return { title: applyPh_(it.title, map), max: parseInt(c[1] || '5', 10), w: it.weight };
  });
  if (!scales.length) { ui.alert('審査フォーム設問に「目盛」の設問がありません。'); return; }
  const texts = q.items.filter(it => it.type === '段落').map(it => applyPh_(it.title, map));
  const nameTitle = (q.items.find(it => it.type === '記述式') || {}).title || '';

  const data = pres.getRange(PRES_FIRST, 1, PRES_MAX, 7).getDisplayValues();
  const results = [];
  const comments = [];
  data.forEach(d => {
    if (!d[1].trim() || !d[6]) return;
    const form = FormApp.openByUrl(d[6]);
    const resps = form.getResponses();
    const sum = scales.map(() => 0), cnt = scales.map(() => 0);
    resps.forEach(r => {
      let reviewer = '';
      const its = r.getItemResponses();
      its.forEach(ir => { if (ir.getItem().getTitle() === applyPh_(nameTitle, map)) reviewer = String(ir.getResponse()); });
      its.forEach(ir => {
        const t = ir.getItem().getTitle();
        const si = scales.findIndex(s => s.title === t);
        if (si >= 0) {
          const v = Number(ir.getResponse());
          if (!isNaN(v)) { sum[si] += v; cnt[si]++; }
        } else if (texts.indexOf(t) >= 0 && String(ir.getResponse()).trim()) {
          comments.push([d[0], d[1], reviewer, t, String(ir.getResponse())]);
        }
      });
    });
    const avgs = scales.map((s, i) => cnt[i] ? sum[i] / cnt[i] : '');
    let num = 0, den = 0;
    scales.forEach((s, i) => { if (avgs[i] !== '') { num += s.w * avgs[i] / s.max; den += s.w; } });
    results.push({
      no: d[0], name: d[1], title: d[2], n: resps.length, avgs: avgs,
      score: den ? Math.round(num / den * 10000) / 100 : '',
    });
  });
  if (!results.length) { ui.alert('集計対象がありません（審査フォーム未作成）。'); return; }

  // 並べ替えと順位（同点は同順位）
  results.sort((a, b) => (b.score === '' ? -1 : b.score) - (a.score === '' ? -1 : a.score));
  let rank = 0, prev = null;
  results.forEach((r, i) => {
    if (r.score === '') { r.rank = '-'; return; }
    if (r.score !== prev) rank = i + 1;
    r.rank = rank; prev = r.score;
  });

  const sh = ss.getSheetByName(SUM_SHEET) || ss.insertSheet(SUM_SHEET);
  sh.clear(); sh.setConditionalFormatRules([]);
  sh.getRange('A1').setValue('審査集計（点数・順位）').setFontSize(14).setFontWeight('bold');
  sh.getRange('A2').setValue('更新: ' + Utilities.formatDate(new Date(), ss.getSpreadsheetTimeZone(), 'yyyy/MM/dd HH:mm') +
    ' ／ 総合得点は目盛項目の平均を最大値で割り、配点の重み（審査フォーム設問シート H列）で加重平均して100点満点に換算').setFontColor('#666666');
  const best = results.filter(r => r.rank === 1);
  if (best.length) sh.getRange('A3').setValue('最優秀賞候補： ' + best.map(r => r.name).join(' / ')).setFontWeight('bold').setFontColor('#c00000');
  const head = ['順位', '番号', '発表者', '論文タイトル', '回答数'].concat(scales.map(s => s.title + '\n(平均/' + s.max + ')'), ['総合得点\n(100点満点)']);
  styleHeader_(sh.getRange(5, 1, 1, head.length).setValues([head])).setWrap(true);
  const rows = results.map(r => [r.rank, r.no, r.name, r.title, r.n].concat(r.avgs.map(a => a === '' ? '' : Math.round(a * 100) / 100), [r.score]));
  sh.getRange(6, 1, rows.length, head.length).setValues(rows).setVerticalAlignment('middle');
  sh.getRange(6, 6, rows.length, scales.length).setNumberFormat('0.00');
  sh.getRange(6, head.length, rows.length, 1).setNumberFormat('0.00').setFontWeight('bold');
  sh.getRange(6, 1, rows.length, 2).setHorizontalAlignment('center');
  const colors = { 1: '#ffe699', 2: '#e7e6e6', 3: '#f8cbad' };
  rows.forEach((r, i) => { if (colors[r[0]]) sh.getRange(6 + i, 1, 1, head.length).setBackground(colors[r[0]]); });
  [50, 50, 200, 280, 60].concat(scales.map(() => 95), [110]).forEach((w, i) => sh.setColumnWidth(i + 1, w));
  sh.setRowHeight(5, 50);
  sh.setFrozenRows(5);

  const cs = ss.getSheetByName(CMT_SHEET) || ss.insertSheet(CMT_SHEET);
  cs.clear();
  styleHeader_(cs.getRange(1, 1, 1, 5).setValues([['番号', '発表者', '審査者', '項目', 'コメント']]));
  if (comments.length) cs.getRange(2, 1, comments.length, 5).setValues(comments).setWrap(true).setVerticalAlignment('top');
  [50, 160, 140, 140, 600].forEach((w, i) => cs.setColumnWidth(i + 1, w));
  cs.setFrozenRows(1);
  ss.setActiveSheet(sh);
  ui.alert('集計しました。「' + SUM_SHEET + '」「' + CMT_SHEET + '」シートを確認してください。');
}
