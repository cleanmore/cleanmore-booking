/* ================================================================
   作業完了のお礼メール＋Googleの口コミ依頼（自動送信）

   ・Apps Script に「ReviewRequest.gs」として追加するだけ（Code.gs は触らない）
   ・毎日19時に「予約データ」を見て、ステータスが「完了」で未送信のお客様に1回だけ送る
   ・予約データ＝予約フォームから来た直接のお客様だけ。sales.html で入れたポータルの仕事には送らない
   ・送ったかどうかは「予約データ」の右端の列「口コミ依頼」に記録（同じ人に二重送信しない）

   初回の準備：setupReviewRequest を1回実行 → 「設定」シートの REVIEW_URL に口コミURLを入れる
   ================================================================ */

const REVIEW = {
  BOOKING_SHEET: '予約データ',
  SETTINGS_SHEET: '設定',
  DONE_STATUS: '完了',
  MARK_HEADER: '口コミ依頼',
  SEND_HOUR: 19,
  // 列は見出しの文字で探す（見出しを変えたらここも合わせる）
  HEADERS: {
    status: ['ステータス'],
    email: ['メールアドレス', 'メール', 'Eメール', 'email'],
    name: ['お名前', '氏名', '名前', 'name']
  }
};

/** 初回に1回だけ実行：設定の追加・毎日の自動実行の登録・今までの完了分を「対象外」にする */
function setupReviewRequest() {
  ensureSetting_('REVIEW_URL', '', 'Googleの口コミ投稿用URL（ビジネスプロフィールの「口コミを依頼」からコピー）');
  ensureSetting_('REVIEW_MAIL_ENABLED', 'はい', '完了後のお礼＋口コミ依頼メールを送るか（はい／いいえ）');

  ScriptApp.getProjectTriggers()
    .filter(t => t.getHandlerFunction() === 'sendReviewRequests')
    .forEach(t => ScriptApp.deleteTrigger(t));
  ScriptApp.newTrigger('sendReviewRequests')
    .timeBased().everyDays(1).atHour(REVIEW.SEND_HOUR).inTimezone('Asia/Tokyo').create();

  // 導入前に完了したお客様へ急に届かないよう、既存の「完了」は送信済み扱いにする
  const t = bookingTable_();
  let skipped = 0;
  t.rows.forEach((r, i) => {
    if (r[t.col.status] === REVIEW.DONE_STATUS && !r[t.col.mark]) {
      t.sheet.getRange(i + 2, t.col.mark + 1).setValue('対象外（導入前）');
      skipped++;
    }
  });

  SpreadsheetApp.getUi().alert(
    '口コミ依頼メールの準備ができました。\n\n' +
    '1. 「設定」シートの REVIEW_URL に Googleの口コミ投稿用URL を貼ってください\n' +
    '2. testReviewMail を実行すると、自分宛てにお試しメールが届きます\n\n' +
    `毎日${REVIEW.SEND_HOUR}時ごろ自動で送ります。（導入前の完了 ${skipped} 件は送りません）`
  );
}

/** 毎日自動で実行される：完了＆未送信のお客様に送る */
function sendReviewRequests() {
  if (getSetting_('REVIEW_MAIL_ENABLED') !== 'はい') return;
  const url = getSetting_('REVIEW_URL');
  if (!url) return;

  const t = bookingTable_();
  t.rows.forEach((r, i) => {
    if (r[t.col.status] !== REVIEW.DONE_STATUS || r[t.col.mark]) return;
    const cell = t.sheet.getRange(i + 2, t.col.mark + 1);
    const email = String(r[t.col.email] || '').trim();
    if (!/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email)) { cell.setValue('メールなし'); return; }
    try {
      const mail = reviewMail_(String(r[t.col.name] || '').trim(), url);
      MailApp.sendEmail({ to: email, subject: mail.subject, body: mail.body, name: 'クリーンモア' });
      cell.setValue(Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy/MM/dd HH:mm') + ' 送信');
    } catch (err) {
      cell.setValue('送信エラー：' + err.message); // 次の日は再送しない。手で消せば再送される
    }
  });
}

/** 自分宛てにお試しメールを送る */
function testReviewMail() {
  const url = getSetting_('REVIEW_URL') || '（REVIEW_URL が未設定です）';
  const mail = reviewMail_('テスト', url);
  MailApp.sendEmail({ to: Session.getActiveUser().getEmail(), subject: '【テスト】' + mail.subject, body: mail.body, name: 'クリーンモア' });
}

/** メールの文面（ここを書き換えれば文面が変わる） */
function reviewMail_(name, url) {
  const tel = getSetting_('SHOP_TEL');
  const to = name ? `${name} 様` : 'お客様';
  return {
    subject: '【クリーンモア】本日はエアコンクリーニングのご依頼ありがとうございました',
    body:
`${to}

本日はクリーンモアにご依頼いただき、ありがとうございました。
エアコンの使い心地はいかがでしょうか。

━━━━━━━━━━━━━━━━━━
▼ きれいな状態を長持ちさせるコツ
・冷房や除湿のあとは「送風」を30分〜1時間
・フィルターは2週間に1回ほど、ホコリを取る
・次回のクリーニングは1年後がおすすめです
━━━━━━━━━━━━━━━━━━

【お願い】
もしよろしければ、今回の感想をGoogleにお寄せいただけないでしょうか。
「どのお部屋のエアコンだったか」「ニオイや効きがどう変わったか」など、
ひとことでも、これからご依頼を考えている長崎の方の参考になります。

▼ 感想を書く（1分ほどで終わります）
${url}

いただいたお声はスタッフ全員で拝見し、励みにしています。
気になることがあれば、このメールへの返信やお電話でいつでもご連絡ください。

━━━━━━━━━━━━━━━━━━
エアコンクリーニング クリーンモア
☎ ${tel}
家も心もスッキリするクリーニングを提供します🌿
━━━━━━━━━━━━━━━━━━`
  };
}

/* ---------- ここから下は道具 ---------- */

function bookingTable_() {
  const sheet = SpreadsheetApp.getActive().getSheetByName(REVIEW.BOOKING_SHEET);
  if (!sheet) throw new Error(`シート「${REVIEW.BOOKING_SHEET}」が見つかりません`);
  let headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0].map(h => String(h).trim());
  if (!headers.includes(REVIEW.MARK_HEADER)) {
    sheet.getRange(1, headers.length + 1).setValue(REVIEW.MARK_HEADER);
    headers.push(REVIEW.MARK_HEADER);
  }
  const find = (names, label) => {
    let i = headers.findIndex(h => names.includes(h));
    // 「メール配信」などを誤って拾わないよう、完全一致がなければ部分一致（配信を除く）
    if (i < 0) i = headers.findIndex(h => names.some(n => h.includes(n)) && !h.includes('配信'));
    if (i < 0) throw new Error(`「${REVIEW.BOOKING_SHEET}」に「${label}」の列が見つかりません（REVIEW.HEADERS を見出しに合わせてください）`);
    return i;
  };
  const col = {
    status: find(REVIEW.HEADERS.status, 'ステータス'),
    email: find(REVIEW.HEADERS.email, 'メール'),
    name: find(REVIEW.HEADERS.name, 'お名前'),
    mark: headers.indexOf(REVIEW.MARK_HEADER)
  };
  const n = sheet.getLastRow() - 1;
  const rows = n > 0 ? sheet.getRange(2, 1, n, headers.length).getValues() : [];
  return { sheet, col, rows };
}

// 「設定」シートは A列＝キー、B列＝値、C列＝説明 の前提
function getSetting_(key) {
  const sheet = SpreadsheetApp.getActive().getSheetByName(REVIEW.SETTINGS_SHEET);
  if (!sheet) return '';
  const row = sheet.getDataRange().getValues().find(r => String(r[0]).trim() === key);
  return row ? String(row[1]).trim() : '';
}

function ensureSetting_(key, value, note) {
  const sheet = SpreadsheetApp.getActive().getSheetByName(REVIEW.SETTINGS_SHEET);
  if (!sheet) throw new Error(`シート「${REVIEW.SETTINGS_SHEET}」が見つかりません`);
  const exists = sheet.getDataRange().getValues().some(r => String(r[0]).trim() === key);
  if (!exists) sheet.appendRow([key, value, note]);
}
