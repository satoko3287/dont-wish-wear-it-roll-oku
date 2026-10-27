// =====================================================
// WEAR IT & ROLL -奥の間- 申込フォーム → Google スプレッドシート
// =====================================================
// 【使い方】
// 1. Google スプレッドシートを新規作成する
// 2. メニュー「拡張機能」→「Apps Script」を開く
// 3. このコードを丸ごとコピー&ペーストして保存
// 4. 「デプロイ」→「新しいデプロイ」をクリック
// 5. 種類：「ウェブアプリ」、アクセス：「全員」に設定してデプロイ
// 6. 発行されたURLを index.html の GAS_URL に貼り付ける
//
// 【LINE通知の設定】（LINE Notify は2025年3月で終了したため Messaging API を使用）
// 1. https://developers.line.biz/ にLINEアカウントでログイン
// 2. プロバイダーを作成 →「Messaging API」チャネルを作成（＝LINE公式アカウントができる）
// 3. チャネルの「Messaging API設定」タブ
//    → 一番下の「チャネルアクセストークン（長期）」を「発行」してコピー
//    → 同じタブのQRコードを自分のLINEで読み取り、友だち追加する
// 4. チャネルの「チャネル基本設定」タブ → 一番下の「あなたのユーザーID」（U から始まる）をコピー
// 5. Apps Script 左メニュー「⚙ プロジェクトの設定」→「スクリプト プロパティ」に以下を追加
//      LINE_CHANNEL_ACCESS_TOKEN = （3でコピーしたトークン）
//      LINE_USER_ID              = （4でコピーしたユーザーID）
// 6. エディタで testLine を選んで ▶実行 → 権限を許可 → LINEに届けばOK
// 7. 「デプロイ」→「デプロイを管理」→ ✏ 編集 → バージョン「新バージョン」→ デプロイ
//    （※これをしないと修正が本番に反映されません。URLは変わりません）
//
// 【メンテナンス申込フォーム（トップページ下部 / LIFF）の設定】
// 1. LINE Developers で、公式アカウントと同じプロバイダーに「LINEログイン」チャネルを作成
//    （アプリタイプ：ウェブアプリ）
// 2. 作成したチャネルの「チャネル基本設定」タブ
//    →「リンクされたLINE公式アカウント」で公式アカウントを選ぶ
//    → 上のほうの「チャネルID」（数字）をコピー
// 3. 同じチャネルの「LIFF」タブ →「追加」
//      サイズ：Full ／ エンドポイントURL：https://satoko3287.github.io/wear-it-and-roll-oku/
//      Scope：openid と profile にチェック ／ 友だち追加オプション：On (Aggressive)
//    → 発行された「LIFF ID」を index.html の LIFF_ID に設定
//    → 「LIFF URL」（https://liff.line.me/...）をQRコードやリンクにしてお客様に渡す
//      （LINEの中で開くと、自動でメンテナンス申込フォームまでスクロールします）
// 4. スクリプト プロパティに以下を追加
//      LINE_LOGIN_CHANNEL_ID = （2でコピーしたチャネルID）
//    ※これがあると、お客様のLINEに「受付しました」のメッセージを自動で送ります
// 5. 申込内容は同じスプレッドシートの「メンテナンス」シートに記録されます（自動作成）
// =====================================================

// ★ 通知メールの送り先（自分のメールアドレスに変更してください）
const NOTIFY_EMAIL = 'life.is.sweet.satoko@gmail.com';

// お客様からのお問い合わせ先（受付メールの返信先・署名に表示）
const CONTACT_EMAIL = '3287minihana@gmail.com';

// お客様に届く受付メールの差出人名と署名
const SHOP_NAME = 'WEAR IT & ROLL -奥の間-';
const MAIL_SIGNATURE = `
──────────────────────
WEAR IT & ROLL -奥の間-
まとって、さあ行こう

運営：合同会社みにはな
お問い合わせ：3287minihana@gmail.com
ホームページ：https://satoko3287.github.io/wear-it-and-roll-oku/
公式LINE：https://line.me/R/ti/p/@546copgu
──────────────────────`;

// お客様へ受付メールを送る（メールアドレスの入力がある場合のみ。返信は CONTACT_EMAIL に届く）
function sendCustomerMail(to, subject, body) {
  if (!to || !/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(to)) return;
  try {
    MailApp.sendEmail({
      to: to,
      subject: subject,
      body: body + '\n' + MAIL_SIGNATURE,
      name: SHOP_NAME,
      replyTo: CONTACT_EMAIL,
    });
  } catch (err) {
    console.error('お客様への受付メール送信エラー: ' + err.message);
  }
}

const HEADERS = [
  '受付日時',
  '旧姓お名前（ひらがな）',
  '生年月日',
  'メールアドレス',
  'プラン',
  'サイズ',
  '目的',
  '石の方向性',
  '送り先お名前',
  '郵便番号',
  '電話番号',
  '住所',
  'お支払い方法',
  '紹介者',
  'LINE名',
  'LINEユーザーID',
];

const MAINTENANCE_SHEET_NAME = 'メンテナンス';

const MAINTENANCE_HEADERS = [
  '受付日時',
  'メニュー',
  'ご希望サイズ',
  'ご要望・状態',
  'ご購入時期',
  'お名前',
  '電話番号',
  'メールアドレス',
  '返送先お名前',
  '郵便番号',
  '住所',
  'お支払い方法',
  'LINE名',
  'LINEユーザーID',
  '紹介者',
];

// 1行目の見出しを用意する（新しい列が増えたときは見出しだけ書き足す）
function ensureHeaders(sheet, headers, background, fontColor) {
  const lastCol = sheet.getLastRow() === 0 ? 0 : sheet.getLastColumn();
  if (lastCol >= headers.length) return;
  sheet.getRange(1, lastCol + 1, 1, headers.length - lastCol)
    .setValues([headers.slice(lastCol)]);
  sheet.getRange(1, 1, 1, headers.length)
    .setFontWeight('bold')
    .setBackground(background)
    .setFontColor(fontColor);
  sheet.setFrozenRows(1);
}

function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);
    if (data.formType === 'maintenance') {
      return handleMaintenance(data);
    }

    // オーダーは1枚目のシートに記録（メンテナンス用シートを追加しても変わらないように）
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheets()[0];

    ensureHeaders(sheet, HEADERS, '#1a1730', '#c9a96e');

    // LINEのIDトークンを検証して、本人のLINEユーザーIDを取り出す（なりすまし防止）
    const lineUser = verifyLineIdToken(data.lineIdToken);

    sheet.appendRow([
      new Date(),
      data.kana     || '',
      data.birthday || '',
      data.email    || '',
      data.plan     || '',
      data.size     || '',
      data.purpose  || '',
      data.stone    || '',
      data.shipName || '',
      data.postal   || '',
      data.phone    || '',
      data.address  || '',
      data.payment  || '',
      data.referrer || '',
      lineUser ? lineUser.name : '',
      lineUser ? lineUser.userId : '',
    ]);

    // ── 通知本文 ──────────────────────────────
    const subject = '【奥の間】新しいオーダーが届きました';
    const body = `
新しいオーダーが届きました！
スプレッドシートをご確認ください。

──────────────────────
■ 受付日時：${new Date().toLocaleString('ja-JP')}
■ お名前（ひらがな）：${data.kana || ''}
■ 生年月日：${data.birthday || ''}
■ メールアドレス：${data.email || ''}
■ プラン：${data.plan || ''}
■ サイズ：${data.size || ''}
■ 目的：${data.purpose || ''}
■ 石の方向性：${data.stone || '（記載なし）'}
■ 紹介者：${data.referrer || '（なし）'}
■ LINE名：${lineUser ? lineUser.name : '（LINE外から申込）'}
──────────────────────
■ 送り先お名前：${data.shipName || ''}
■ 郵便番号：${data.postal || ''}
■ 住所：${data.address || ''}
■ 電話番号：${data.phone || ''}
■ お支払い方法：${data.payment || ''}
──────────────────────

-奥の間- 自動通知
`.trim();

    // ── LINE通知（失敗してもメール通知・スプレッドシート記録は続行）──
    try {
      sendLine(subject + '\n\n' + body);
    } catch (lineErr) {
      console.error('LINE通知エラー: ' + lineErr.message);
    }

    // ── メール通知 ──
    try {
      MailApp.sendEmail(NOTIFY_EMAIL, subject, body);
    } catch (mailErr) {
      console.error('メール通知エラー: ' + mailErr.message);
    }

    // ── お客様のLINEへ受付メッセージ ──
    if (lineUser) {
      try {
        pushLine(lineUser.userId, `
${data.shipName || lineUser.name} 様

オーダーのお申し込みありがとうございます。
以下の内容で受け付けました。

■ プラン：${data.plan || ''}
■ サイズ：${data.size || ''}
■ お支払い方法：${data.payment || ''}

24時間以内に、このLINEにてご連絡いたします。
数秘リーディングの詳細やご要望を、こちらのトークで丁寧にお伺いします。

WEAR IT & ROLL -奥の間-
`.trim());
      } catch (pushErr) {
        console.error('お客様への受付メッセージ送信エラー: ' + pushErr.message);
      }
    }

    // ── お客様へ受付メール ──
    sendCustomerMail(data.email, `【${SHOP_NAME}】オーダーのお申し込みを受け付けました`, `
${data.shipName || data.kana || ''} 様

この度は、${SHOP_NAME} にオーダーのお申し込みをいただき、誠にありがとうございます。
以下の内容で受け付けました。

■ プラン：${data.plan || ''}
■ サイズ：${data.size || ''}
■ 目的：${data.purpose || ''}
■ 石の方向性：${data.stone || '（記載なし）'}
■ お届け先：〒${data.postal || ''} ${data.address || ''}　${data.shipName || ''} 様
■ お支払い方法：${data.payment || ''}

24時間以内に、${lineUser ? '公式LINE' : 'このメールアドレス宛て'}にご連絡いたします。
数秘リーディングの詳細やご要望を、丁寧にお伺いします。

ご不明点がございましたら、このメールにご返信ください。
※お心当たりのない場合は、お手数ですがこのメールを破棄してください。
`.trim());

    return ContentService
      .createTextOutput(JSON.stringify({ status: 'success' }))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ status: 'error', message: err.message }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

// ── メンテナンス申込 ──────────────────────────────
function handleMaintenance(data) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(MAINTENANCE_SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(MAINTENANCE_SHEET_NAME, ss.getNumSheets());
  }
  ensureHeaders(sheet, MAINTENANCE_HEADERS, '#1f3460', '#ffffff');

  // LINEのIDトークンを検証して、本人のLINEユーザーIDを取り出す（なりすまし防止）
  const lineUser = verifyLineIdToken(data.lineIdToken);

  sheet.appendRow([
    new Date(),
    data.menu     || '',
    data.wishSize || '',
    data.detail   || '',
    data.purchase || '',
    data.name     || '',
    data.phone    || '',
    data.email    || '',
    data.shipName || '',
    data.postal   || '',
    data.address  || '',
    data.payment  || '',
    lineUser ? lineUser.name : '',
    lineUser ? lineUser.userId : '',
    data.referrer || '',
  ]);

  const subject = '【メンテナンス申込】新しいお申し込みが届きました';
  const body = `
■ 受付日時：${new Date().toLocaleString('ja-JP')}
■ メニュー：${data.menu || ''}${data.wishSize ? '\n■ ご希望サイズ：' + data.wishSize : ''}
■ ご要望・状態：${data.detail || '（記載なし）'}
■ ご購入時期：${data.purchase || '（記載なし）'}
■ 紹介者：${data.referrer || '（なし）'}
──────────────────────
■ お名前：${data.name || ''}
■ 電話番号：${data.phone || ''}
■ メール：${data.email || '（記載なし）'}
■ LINE名：${lineUser ? lineUser.name : '（LINE外から申込）'}
──────────────────────
■ 返送先：${data.shipName || ''}
〒${data.postal || ''}
${data.address || ''}
■ お支払い方法：${data.payment || ''}
`.trim();

  // 自分への通知（LINE・メール）
  try {
    sendLine(subject + '\n\n' + body);
  } catch (lineErr) {
    console.error('LINE通知エラー: ' + lineErr.message);
  }
  try {
    MailApp.sendEmail(NOTIFY_EMAIL, subject, body);
  } catch (mailErr) {
    console.error('メール通知エラー: ' + mailErr.message);
  }

  // お客様のLINEへ受付メッセージ
  if (lineUser) {
    try {
      pushLine(lineUser.userId, `
${data.name || lineUser.name} 様

メンテナンスのお申し込みありがとうございます。
以下の内容で受け付けました。

■ メニュー：${data.menu || ''}${data.wishSize ? '\n■ ご希望サイズ：' + data.wishSize : ''}
■ お支払い方法：${data.payment || ''}

【お送りいただく方法】
ブレスレットをレターパック（青）でお送りください。
中に返送用のレターパック（赤 または 青）を入れてお送りください。

送り先は、このLINEで改めてご案内いたします。
ご不明点がございましたら、このトークでお気軽にお問い合わせください。

WEAR IT & ROLL -奥の間-
`.trim());
    } catch (pushErr) {
      console.error('お客様への受付メッセージ送信エラー: ' + pushErr.message);
    }
  }

  // お客様へ受付メール
  sendCustomerMail(data.email, `【${SHOP_NAME}】メンテナンスのお申し込みを受け付けました`, `
${data.name || ''} 様

この度は、${SHOP_NAME} にメンテナンスのお申し込みをいただき、誠にありがとうございます。
以下の内容で受け付けました。

■ メニュー：${data.menu || ''}${data.wishSize ? '\n■ ご希望サイズ：' + data.wishSize : ''}
■ ご要望・状態：${data.detail || '（記載なし）'}
■ 返送先：〒${data.postal || ''} ${data.address || ''}　${data.shipName || ''} 様
■ お支払い方法：${data.payment || ''}

【お送りいただく方法】
ブレスレットをレターパック（青）でお送りください。
中に返送用のレターパック（赤 または 青）を入れてお送りください。

送り先は、${lineUser ? '公式LINEにて' : 'このメールアドレス宛てに'}改めてご案内いたします。

ご不明点がございましたら、このメールにご返信ください。
※お心当たりのない場合は、お手数ですがこのメールを破棄してください。
`.trim());

  return ContentService
    .createTextOutput(JSON.stringify({ status: 'success' }))
    .setMimeType(ContentService.MimeType.JSON);
}

// LIFFのIDトークンをLINEに問い合わせて検証し、{ userId, name } を返す（無効なら null）
function verifyLineIdToken(idToken) {
  const channelId = PropertiesService.getScriptProperties().getProperty('LINE_LOGIN_CHANNEL_ID');
  if (!idToken || !channelId) return null;

  const res = UrlFetchApp.fetch('https://api.line.me/oauth2/v2.1/verify', {
    method: 'post',
    payload: { id_token: idToken, client_id: channelId },
    muteHttpExceptions: true,
  });
  if (res.getResponseCode() !== 200) {
    console.error('IDトークン検証エラー: ' + res.getContentText());
    return null;
  }
  const json = JSON.parse(res.getContentText());
  return { userId: json.sub, name: json.name || '' };
}

// 自分（LINE_USER_ID）にLINEを送信
function sendLine(text) {
  const userId = PropertiesService.getScriptProperties().getProperty('LINE_USER_ID');
  if (!userId) {
    throw new Error('スクリプトプロパティ LINE_USER_ID が未設定です');
  }
  pushLine(userId, text);
}

// LINE Messaging API で指定したユーザーにプッシュメッセージを送信
function pushLine(to, text) {
  const token = PropertiesService.getScriptProperties().getProperty('LINE_CHANNEL_ACCESS_TOKEN');
  if (!token) {
    throw new Error('スクリプトプロパティ LINE_CHANNEL_ACCESS_TOKEN が未設定です');
  }

  const res = UrlFetchApp.fetch('https://api.line.me/v2/bot/message/push', {
    method: 'post',
    contentType: 'application/json',
    headers: { Authorization: 'Bearer ' + token },
    payload: JSON.stringify({
      to: to,
      // LINEの1メッセージ上限は5000文字
      messages: [{ type: 'text', text: text.slice(0, 5000) }],
    }),
    muteHttpExceptions: true,
  });

  const code = res.getResponseCode();
  if (code !== 200) {
    throw new Error('LINE API ' + code + ': ' + res.getContentText());
  }
}

// LINE通知のテスト（エディタで testLine を選んで ▶実行）
function testLine() {
  sendLine('【奥の間】LINE通知のテストです。これが届いていれば設定完了です✨');
  Logger.log('LINE送信OK');
}

// テスト用（Apps Script エディタから手動実行して動作確認できます）
function testPost() {
  const dummy = {
    postData: {
      contents: JSON.stringify({
        kana:     'やまだ はなこ',
        birthday: '1990-01-15',
        email:    'test@example.com',
        plan:     'テストプラン',
        size:     '内径約 16cm',
        purpose:  '日常のお守り・エネルギー調整として',
        stone:    '落ち着いた青系が好きです',
        shipName: '山田 花子',
        postal:   '000-0000',
        phone:    '090-0000-0000',
        address:  '東京都渋谷区〇〇 1-2-3',
        payment:  '銀行振込',
      })
    }
  };
  const result = doPost(dummy);
  Logger.log(result.getContent());
}

// メンテナンス申込のテスト（エディタで testMaintenance を選んで ▶実行）
// →「メンテナンス」シートに1行追加され、自分のLINEに通知が届けばOK
function testMaintenance() {
  const dummy = {
    postData: {
      contents: JSON.stringify({
        formType: 'maintenance',
        menu:     'ゴム交換（浄化付き）¥3,000',
        detail:   'テスト送信です',
        name:     'テスト 花子',
        phone:    '090-0000-0000',
        shipName: 'テスト 花子',
        postal:   '000-0000',
        address:  '東京都渋谷区〇〇 1-2-3',
        payment:  '銀行振込',
      })
    }
  };
  Logger.log(doPost(dummy).getContent());
}
