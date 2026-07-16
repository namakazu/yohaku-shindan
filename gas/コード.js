// ============================================
// 余白診断 回答ログ 受け口（doPost）
// スプレッドシート「余白診断_回答ログ」のコンテナバインドで使用
// ============================================

// 実際の余白診断の6軸キー名（診断HTML内のAXES定義と一致させている）
// 順番 = シートB〜G列の順（時間・判断・感情・通知・空間・お金）
const SCORE_KEYS = ['time', 'mind', 'heart', 'notify', 'space', 'money'];

// resultTypeの許可リスト（6タイプ + 「余白、保てています。」＝keep）
const ALLOWED_RESULT_TYPES = SCORE_KEYS.concat(['keep']);
const ALLOWED_AGE_GROUPS = ['20s_under', '30s', '40s', '50s_over'];
const ALLOWED_OCCUPATIONS = ['employee', 'freelance', 'homemaker', 'student', 'other'];

const LOG_SHEET_NAME = 'log';

// URLはHTMLソースから誰でも読めるため、この検証・レート制限は
// あくまで雑なゴミデータ・簡易フラッド対策であり、悪意ある専用スクリプトは防げない前提
const RATE_LIMIT_MAX = 30;       // 1ウィンドウあたりの最大書き込み数
const RATE_LIMIT_WINDOW_SEC = 60;

function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);

    // 想定外のデータ（いたずらPOST等）はresultTypeの許可リストで足切り
    if (!data || ALLOWED_RESULT_TYPES.indexOf(data.resultType) === -1) {
      return ContentService.createTextOutput('ng');
    }

    if (isRateLimited()) {
      return ContentService.createTextOutput('rate_limited');
    }

    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(LOG_SHEET_NAME);

    // スコア6軸をキー順に展開（0-9の数値以外は不正値として空欄扱い）
    const scores = SCORE_KEYS.map(function (key) {
      const v = data.scores && data.scores[key];
      return isValidScore(v) ? v : '';
    });

    const attr = data.attributes || {};

    sheet.appendRow([
      new Date(),                                  // timestamp（サーバー時刻で記録）
      scores[0], scores[1], scores[2],
      scores[3], scores[4], scores[5],
      data.resultType,
      ALLOWED_AGE_GROUPS.indexOf(attr.ageGroup) !== -1 ? attr.ageGroup : '',
      ALLOWED_OCCUPATIONS.indexOf(attr.occupation) !== -1 ? attr.occupation : '',
      typeof attr.hasChild === 'boolean' ? attr.hasChild : '',
      typeof data.referrer === 'string' ? data.referrer.slice(0, 100) : 'direct'
    ]);

    return ContentService.createTextOutput('ok');
  } catch (err) {
    // 受け口側のエラーで診断体験に影響は出ないが、原因調査用に記録しておく
    console.error('doPost error: ' + err);
    return ContentService.createTextOutput('error');
  }
}

function isValidScore(v) {
  return typeof v === 'number' && isFinite(v) && v >= 0 && v <= 9;
}

// 匿名Webアプリのため送信者を識別できず、全体共通の粗いフラッド対策にしかならない
function isRateLimited() {
  const cache = CacheService.getScriptCache();
  const bucket = 'rl_' + Math.floor(Date.now() / (RATE_LIMIT_WINDOW_SEC * 1000));
  const current = Number(cache.get(bucket) || '0');
  if (current >= RATE_LIMIT_MAX) return true;
  cache.put(bucket, String(current + 1), RATE_LIMIT_WINDOW_SEC + 5);
  return false;
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
        resultType: 'time', // 許可リスト対象の値でなければ'ng'扱いになるため実在のresultTypeを使う
        attributes: { ageGroup: '30s', occupation: 'employee', hasChild: true },
        referrer: 'direct'
      })
    }
  };
  const result = doPost(fakeEvent);
  console.log('結果: ' + result.getContent()); // 「ok」が出て、logシートに1行増えていれば成功
}