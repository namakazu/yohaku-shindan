// ============================================
// 余白診断 回答ログ 受け口（doPost）
// スプレッドシート「余白診断_回答ログ」のコンテナバインドで使用
// ============================================

// 実際の余白診断の6軸キー名（診断HTML内のAXES定義と一致させている）
// 順番 = シートB〜G列の順（時間・判断・感情・通知・空間・お金）
const SCORE_KEYS = ['time', 'mind', 'heart', 'notify', 'space', 'money'];

const LOG_SHEET_NAME = 'log';

function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);

    // 想定外のデータ（いたずらPOST等）は resultType の有無で足切り
    if (!data || !data.resultType) {
      return ContentService.createTextOutput('ng');
    }

    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(LOG_SHEET_NAME);

    // スコア6軸をキー順に展開（欠けていても空欄で埋まるだけで落ちない）
    const scores = SCORE_KEYS.map(function (key) {
      return (data.scores && data.scores[key] != null) ? data.scores[key] : '';
    });

    const attr = data.attributes || {};

    sheet.appendRow([
      new Date(),                                  // timestamp（サーバー時刻で記録）
      scores[0], scores[1], scores[2],
      scores[3], scores[4], scores[5],
      data.resultType,
      attr.ageGroup != null ? attr.ageGroup : '',   // 未回答nullは空欄で記録
      attr.occupation != null ? attr.occupation : '',
      attr.hasChild != null ? attr.hasChild : '',
      data.referrer || 'direct'
    ]);

    return ContentService.createTextOutput('ok');
  } catch (err) {
    // 受け口側のエラーで診断体験に影響は出ないが、原因調査用に記録しておく
    console.error('doPost error: ' + err);
    return ContentService.createTextOutput('error');
  }
}

// ============================================
// テスト用：doPostはエディタから直接実行できないため、
// この関数を実行して疑似データで書き込み確認する
// ============================================
function testDoPost() {
  const fakeEvent = {
    postData: {
      contents: JSON.stringify({
        timestamp: new Date().toISOString(),
        scores: { time: 3, mind: 5, heart: 2, notify: 4, space: 1, money: 5 },
        resultType: 'test_type',
        attributes: { ageGroup: '30s', occupation: 'employee', hasChild: true },
        referrer: 'direct'
      })
    }
  };
  const result = doPost(fakeEvent);
  console.log('結果: ' + result.getContent()); // 「ok」が出て、logシートに1行増えていれば成功
}