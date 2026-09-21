/**
 * スマレジ・タイムカード → 管理者本人のLINE通知（Google Apps Script / V8）
 * 必須設定は「プロジェクトの設定」→「スクリプト プロパティ」に入力します。
 * 認証情報をコード・実行ログ・チャットに貼り付ける必要はありません。
 * 利用順: sendTestMessage → checkConnections → startNotifications
 * スマレジに対して行う操作は勤怠実績の読み取りだけです。
 */
const SMN = Object.freeze({
  timezone: 'Asia/Tokyo',
  pollMinutes: 5,
  lookbackDays: 3,
  maxAgeHours: 48,
  retentionDays: 7,
  pageSize: 200,
  maxPages: 20,
  maxSendsPerRun: 20,
  retryHours: 23,
  prefix: 'SMN_',
  donePrefix: 'SMN_DONE_',
  outboxKey: 'SMN_OUTBOX',
  handler: 'pollTimecard'
});

/** LINEの設定だけで実行できます。同じ設定でのテスト送信は1回です。 */
function sendTestMessage() {
  return locked_(function () {
    const c = lineConfig_();
    const p = props_();
    const profile = lineGet_(c, '/profile/' + c.userId);
    const key = SMN.donePrefix + digest_('TEST|' + lineIdentity_(c));
    if (p.getProperty(key)) {
      console.log('この設定でのテストは送信済みです。LINEのトークを確認してください。');
      return;
    }
    const existing = readOutbox_();
    if (existing && existing.key !== key) {
      throw new Error('勤怠通知の送信待ちがあります。先に pollTimecard を実行してください。');
    }
    if (!existing) {
      writeOutbox_(c, key, 'スマレジ・タイムカードのLINE通知テストです。\nこのメッセージが届けば、LINE側の接続は完了です。');
    }
    sendOutbox_(c);
    console.log('テストをLINEへ送信しました。宛先の表示名: ' + clean_(profile.displayName));
  });
}

/** 勤怠と通知先を確認します。LINEメッセージは送りません。 */
function checkConnections() {
  return locked_(function () {
    const c = config_();
    token_(c, true);
    const profile = lineGet_(c, '/profile/' + c.userId);
    const rows = fetchResults_(c, new Date());
    events_(rows, c); // レスポンス形式・時刻を検証。位置情報等は保存しません。
    const quota = quota_(c);
    props_().setProperty('SMN_CHECKED', configIdentity_(c));
    console.log('通知先のLINE表示名: ' + clean_(profile.displayName));
    console.log('直近3日分の勤怠取得: ' + rows.length + '件');
    console.log('LINE当月の使用通数: ' + quota.used + ' / ' + quota.limit);
    console.log('LINE表示名が自分であることと、テスト通知が届いていることを確認してください。');
  });
}

/** 初回は過去の記録を通知済みとして登録し、5分間隔で通知を開始します。 */
function startNotifications() {
  return locked_(function () {
    const c = config_();
    const p = props_();
    const identity = configIdentity_(c);
    if (p.getProperty('SMN_CHECKED') !== identity) {
      throw new Error('最初に checkConnections を実行してください。');
    }
    if (!p.getProperty(SMN.donePrefix + digest_('TEST|' + lineIdentity_(c)))) {
      throw new Error('最初に sendTestMessage を実行して通知先を確認してください。');
    }
    if (p.getProperty('SMN_IDENTITY') && p.getProperty('SMN_IDENTITY') !== identity) {
      throw new Error('対象契約・通知先・対象スタッフの設定が変わっています。別のGASプロジェクトで設定してください。');
    }
    if (!p.getProperty('SMN_STARTED_AT')) {
      const snapshot = events_(fetchResults_(c, new Date()), c);
      const now = Date.now();
      const baseline = {};
      snapshot.filter(e => e.at <= now).forEach(e => { baseline[e.key] = String(now); });
      if (Object.keys(baseline).length) p.setProperties(baseline, false);
      p.setProperties({'SMN_STARTED_AT': String(now), 'SMN_IDENTITY': identity}, false);
    }
    const current = ScriptApp.getProjectTriggers().filter(t => t.getHandlerFunction() === SMN.handler);
    if (current.length === 0) ScriptApp.newTrigger(SMN.handler).timeBased().everyMinutes(SMN.pollMinutes).create();
    current.slice(1).forEach(t => ScriptApp.deleteTrigger(t));
    p.setProperty('SMN_ENABLED', '1');
    console.log('5分間隔の出勤・退勤通知を開始しました。');
  });
}

/** 自動実行する関数です。手動で実行しても同じ打刻は再送しません。 */
function pollTimecard() {
  return locked_(function () {
    const p = props_();
    if (p.getProperty('SMN_ENABLED') !== '1') return;
    try {
      const c = config_();
      if (p.getProperty('SMN_IDENTITY') !== configIdentity_(c)) {
        throw new Error('通知開始時から対象設定が変更されています。stopNotifications で停止し、設定を確認してください。');
      }
      if (readOutbox_()) sendOutbox_(c);
      const now = Date.now();
      const allEvents = events_(fetchResults_(c, new Date(now)), c);
      const known = p.getProperties();
      const started = Number(p.getProperty('SMN_STARTED_AT'));
      const cutoff = Math.max(started, now - SMN.maxAgeHours * 3600000);
      let sent = 0;
      for (const e of allEvents) {
        if (known[e.key] || e.at > now) continue;
        if (e.at < cutoff) {
          p.setProperty(e.key, String(now));
          known[e.key] = String(now);
          continue;
        }
        if (sent >= SMN.maxSendsPerRun) break;
        writeOutbox_(c, e.key, e.text);
        sendOutbox_(c);
        known[e.key] = String(now);
        sent++;
      }
      cleanup_(now);
      p.setProperty('SMN_LAST_SUCCESS', new Date().toISOString());
      p.deleteProperty('SMN_LAST_ERROR');
      console.log('確認完了。新規通知: ' + sent + '件');
    } catch (e) {
      p.setProperty('SMN_LAST_ERROR', new Date().toISOString() + ' ' + String(e.message).slice(0, 300));
      throw e; // GASの実行履歴・トリガーエラー通知で確認できます。
    }
  });
}

/** 通知を止めます。送信履歴は残し、再開時の二重送信を防ぎます。 */
function stopNotifications() {
  return locked_(function () {
    props_().setProperty('SMN_ENABLED', '0');
    ScriptApp.getProjectTriggers().filter(t => t.getHandlerFunction() === SMN.handler)
      .forEach(t => ScriptApp.deleteTrigger(t));
    console.log('通知を停止しました。再開は startNotifications です。');
  });
}

/** 状態だけを表示します。キーや打刻明細は出力しません。 */
function showStatus() {
  const p = props_();
  console.log('通知: ' + (p.getProperty('SMN_ENABLED') === '1' ? '有効' : '停止'));
  console.log('最終確認成功: ' + (p.getProperty('SMN_LAST_SUCCESS') || '未実行'));
  console.log('送信待ち: ' + (p.getProperty(SMN.outboxKey) ? 'あり' : 'なし'));
  console.log('最終エラー: ' + (p.getProperty('SMN_LAST_ERROR') || 'なし'));
}

/**
 * 障害後の手動復旧専用です。LINEのトークで内容を確認してから実行してください。
 * 送信待ち1件を再送せず処理済みにします。未着の場合はその通知が欠けます。
 */
function skipPendingAfterReview() {
  return locked_(function () {
    const pending = readOutbox_();
    if (!pending) { console.log('送信待ちはありません。'); return; }
    props_().setProperty(pending.key, String(Date.now()));
    props_().deleteProperty(SMN.outboxKey);
    console.log('送信待ち1件を再送せず処理済みにしました。');
  });
}

function props_() { return PropertiesService.getScriptProperties(); }

function locked_(fn) {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(1000)) { console.log('別の実行が進行中です。'); return; }
  try { return fn(); } finally { lock.releaseLock(); }
}

function required_(name) {
  const value = (props_().getProperty(name) || '').trim();
  if (!value) throw new Error('スクリプト プロパティに ' + name + ' を入力してください。');
  return value;
}

function lineConfig_() {
  const c = {lineToken: required_('LINE_ACCESS_TOKEN'), userId: required_('LINE_USER_ID')};
  if (!/^U[0-9a-f]{32}$/.test(c.userId)) throw new Error('LINE_USER_ID は「あなたのユーザーID」の U で始まる33文字です。');
  return c;
}

function config_() {
  const c = lineConfig_();
  c.contract = required_('SMAREGI_CONTRACT_ID');
  c.clientId = required_('SMAREGI_CLIENT_ID');
  c.clientSecret = required_('SMAREGI_CLIENT_SECRET');
  if (!/^[A-Za-z0-9_-]+$/.test(c.contract) || c.contract.startsWith('sb_')) {
    throw new Error('SMAREGI_CONTRACT_ID に本番環境の契約IDを入力してください。');
  }
  c.staffIds = ids_('STAFF_IDS');
  c.storeIds = ids_('STORE_IDS');
  return c;
}

function ids_(name) {
  const raw = (props_().getProperty(name) || '').trim();
  if (!raw) return '';
  const ids = raw.split(',').map(s => s.trim());
  if (ids.some(s => !/^\d+$/.test(s))) throw new Error(name + ' は数字のIDを半角カンマで区切って入力してください。');
  return Array.from(new Set(ids)).sort().join(',');
}

function configIdentity_(c) { return digest_([c.contract, c.userId, c.staffIds, c.storeIds].join('|')); }
function lineIdentity_(c) { return digest_(c.userId + '|' + c.lineToken); }
function digest_(s) {
  return Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, s, Utilities.Charset.UTF_8)
    .map(n => ('0' + ((n + 256) % 256).toString(16)).slice(-2)).join('');
}
function clean_(s) { return String(s || '').replace(/[\r\n\t]/g, ' ').slice(0, 100); }

function token_(c, refresh) {
  const cache = CacheService.getScriptCache();
  const key = 'smn-token-' + digest_([c.contract, c.clientId, c.clientSecret].join('|'));
  if (!refresh) { const cached = cache.get(key); if (cached) return cached; }
  const data = fetchJson_('スマレジ認証', 'https://id.smaregi.jp/app/' + encodeURIComponent(c.contract) + '/token', {
    method: 'post',
    contentType: 'application/x-www-form-urlencoded',
    headers: {Authorization: 'Basic ' + Utilities.base64Encode(c.clientId + ':' + c.clientSecret, Utilities.Charset.UTF_8)},
    payload: {grant_type: 'client_credentials', scope: 'timecard.shifts:read'}
  });
  if (!data.access_token) throw new Error('スマレジのアクセストークンを取得できませんでした。');
  cache.put(key, data.access_token, Math.max(1, Math.min(21600, Number(data.expires_in || 3600) - 60)));
  return data.access_token;
}

function fetchResults_(c, now) {
  const query = {
    from_date: Utilities.formatDate(new Date(now.getTime() - (SMN.lookbackDays - 1) * 86400000), SMN.timezone, 'yyyy-MM-dd'),
    to_date: Utilities.formatDate(now, SMN.timezone, 'yyyy-MM-dd'),
    limit: SMN.pageSize, sort: 'shift_date,staff_id,attendance_at'
  };
  if (c.staffIds) query.staff_id = c.staffIds;
  if (c.storeIds) query.store_id = c.storeIds;
  let token = token_(c, false);
  let combined = [];
  for (let page = 1; page <= SMN.maxPages; page++) {
    query.page = page;
    const qs = Object.keys(query).map(k => encodeURIComponent(k) + '=' + encodeURIComponent(query[k])).join('&');
    const url = 'https://api.smaregi.jp/' + encodeURIComponent(c.contract) + '/timecard/shifts/results?' + qs;
    let response = fetchSafe_('スマレジ勤怠取得', url, {method: 'get', headers: {Authorization: 'Bearer ' + token}});
    if (response.getResponseCode() === 401) {
      token = token_(c, true);
      response = fetchSafe_('スマレジ勤怠取得', url, {method: 'get', headers: {Authorization: 'Bearer ' + token}});
    }
    const data = decodeJson_('スマレジ勤怠取得', response);
    if (!data || !Array.isArray(data.results) || !Number.isInteger(data.pageCount) || data.pageCount < 0 ||
        !Number.isInteger(data.page) || data.page !== page || (data.pageCount === 0 && data.results.length)) {
      throw new Error('勤怠APIの応答形式が想定と異なります。通知を止めて仕様を確認してください。');
    }
    if (data.pageCount > SMN.maxPages) throw new Error('取得件数が多すぎます。STAFF_IDS または STORE_IDS で対象を絞ってください。');
    combined = combined.concat(data.results);
    if (page >= data.pageCount) return combined;
  }
  throw new Error('勤怠の全ページを取得できませんでした。');
}

function events_(rows, c) {
  const unique = {};
  for (const r of rows) {
    if (!r || r.staffId == null || r.storeId == null || !r.attendanceAt || !Object.prototype.hasOwnProperty.call(r, 'leavingAt')) {
      throw new Error('勤怠APIの必須項目が不足しています。通知を止めて仕様を確認してください。');
    }
    if (c.staffIds && !c.staffIds.split(',').includes(String(r.staffId))) continue;
    if (c.storeIds && !c.storeIds.split(',').includes(String(r.storeId))) continue;
    [['attendanceAt', '出勤'], ['leavingAt', '退勤']].forEach(function (pair) {
      const value = r[pair[0]];
      if (value == null || value === '') return;
      if (typeof value !== 'string' || !/T\d{2}:\d{2}:\d{2}(?:\.\d+)?(?:Z|[+-]\d{2}:\d{2})$/.test(value)) {
        throw new Error('勤怠の日時形式が想定と異なります。');
      }
      const at = Date.parse(String(value).replace(/Z$/, '+09:00'));
      if (!Number.isFinite(at)) throw new Error('勤怠の日時を読み取れませんでした。');
      const key = SMN.donePrefix + digest_([c.contract, r.staffId, r.storeId, pair[0], at].join('|'));
      const time = Utilities.formatDate(new Date(at), SMN.timezone, 'M/d HH:mm:ss');
      unique[key] = {key: key, at: at, text: '【' + pair[1] + '】' + clean_(r.staffName || '従業員ID ' + r.staffId) + '\n' +
        time + (r.storeName ? '\n' + clean_(r.storeName) : '')};
    });
  }
  return Object.keys(unique).map(k => unique[k]).sort((a, b) => a.at - b.at || a.key.localeCompare(b.key));
}

function readOutbox_() {
  const raw = props_().getProperty(SMN.outboxKey);
  return raw ? JSON.parse(raw) : null;
}

function writeOutbox_(c, key, text) {
  if (readOutbox_()) throw new Error('前の通知が送信待ちです。上書きせず再送を先に行ってください。');
  props_().setProperty(SMN.outboxKey, JSON.stringify({
    key: key, to: c.userId, identity: lineIdentity_(c), text: text,
    retryKey: Utilities.getUuid(), firstAttemptAt: null
  }));
}

function quota_(c) {
  const limit = lineGet_(c, '/message/quota');
  const usage = lineGet_(c, '/message/quota/consumption');
  if (!Number.isFinite(usage.totalUsage) || (limit.type !== 'limited' && limit.type !== 'none') ||
      (limit.type === 'limited' && !Number.isFinite(limit.value))) {
    throw new Error('LINEの利用通数を取得できませんでした。');
  }
  return {used: usage.totalUsage, limit: limit.type === 'limited' ? limit.value : Infinity};
}

function sendOutbox_(c) {
  const p = props_();
  const q = readOutbox_();
  if (!q) return;
  const hadPriorAttempt = q.firstAttemptAt != null;
  if (q.to !== c.userId || q.identity !== lineIdentity_(c)) throw new Error('送信待ちの作成後にLINEの設定が変わっています。送信履歴を確認してください。');
  if (p.getProperty(q.key)) { p.deleteProperty(SMN.outboxKey); return; }
  if (q.firstAttemptAt != null && Date.now() - q.firstAttemptAt >= SMN.retryHours * 3600000) {
    throw new Error('送信結果不明のまま23時間を超えました。LINEのトークで確認後、skipPendingAfterReview でこの1件を処理済みにしてください。');
  }
  // 一度でも送信したものは、上限到達時も同じキーで結果を確認します。
  // LINE側で受理済みなら409となり、追加の通数を消費しません。
  if (q.firstAttemptAt == null) {
    const quota = quota_(c);
    if (quota.used >= quota.limit) throw new Error('LINEの当月通数上限に達しました。通知は送信待ちです。');
    q.firstAttemptAt = Date.now();
    p.setProperty(SMN.outboxKey, JSON.stringify(q));
  }
  const response = fetchSafe_('LINE送信', 'https://api.line.me/v2/bot/message/push', {
    method: 'post', contentType: 'application/json',
    headers: {Authorization: 'Bearer ' + c.lineToken, 'X-Line-Retry-Key': q.retryKey},
    payload: JSON.stringify({to: q.to, messages: [{type: 'text', text: q.text}]})
  });
  const status = response.getResponseCode();
  const headers = response.getAllHeaders();
  const accepted = Object.keys(headers).some(k => k.toLowerCase() === 'x-line-accepted-request-id' && headers[k]);
  if ((status >= 200 && status < 300) || (status === 409 && accepted)) {
    p.setProperty(q.key, String(Date.now()));
    p.deleteProperty(SMN.outboxKey);
    return;
  }
  if (!hadPriorAttempt && status >= 400 && status < 500 && status !== 409) {
    // 認証・通数などで明確に拒否されたリクエスト。原因解消後は新しいキーで送ります。
    q.firstAttemptAt = null;
    q.retryKey = Utilities.getUuid();
    p.setProperty(SMN.outboxKey, JSON.stringify(q));
  }
  throw httpError_('LINE送信', status);
}

function lineGet_(c, path) {
  return fetchJson_('LINE確認', 'https://api.line.me/v2/bot' + path, {
    method: 'get', headers: {Authorization: 'Bearer ' + c.lineToken}
  });
}

function fetchSafe_(label, url, options) {
  try { return UrlFetchApp.fetch(url, Object.assign({}, options, {muteHttpExceptions: true, followRedirects: false})); }
  catch (_) { throw new Error(label + 'の通信に失敗しました。送信待ちは保持されます。'); }
}
function fetchJson_(label, url, options) { return decodeJson_(label, fetchSafe_(label, url, options)); }
function decodeJson_(label, response) {
  const status = response.getResponseCode();
  if (status < 200 || status >= 300) throw httpError_(label, status);
  try { return JSON.parse(response.getContentText()); }
  catch (_) { throw new Error(label + 'の応答を読み取れませんでした。'); }
}
function httpError_(label, status) {
  const hints = {400: '設定値・対象契約・アプリの利用登録を確認してください。',
    401: '本番用の認証情報・アクセストークンを確認してください。',
    403: 'アプリの利用登録と勤怠参照権限を確認してください。',
    404: '契約ID・LINEユーザーID・友だち追加状態を確認してください。',
    429: 'LINEの通数上限またはAPIの利用制限を確認してください。'};
  return new Error(label + ' HTTP ' + status + '。' + (hints[status] || 'サービス状況と実行履歴を確認してください。'));
}

function cleanup_(now) {
  const p = props_();
  const today = Utilities.formatDate(new Date(now), SMN.timezone, 'yyyy-MM-dd');
  if (p.getProperty('SMN_CLEANED_DATE') === today) return;
  const values = p.getProperties();
  const testKey = SMN.donePrefix + digest_('TEST|' + lineIdentity_(lineConfig_()));
  const cutoff = now - SMN.retentionDays * 86400000;
  Object.keys(values).filter(k => k.startsWith(SMN.donePrefix) && k !== testKey && Number(values[k]) < cutoff)
    .forEach(k => p.deleteProperty(k));
  p.setProperty('SMN_CLEANED_DATE', today);
}
function diagnoseTimecard() {
  return locked_(function () {
    const c = config_();
    const state = props_().getProperties();
    const now = Date.now();
    const started = Number(state.SMN_STARTED_AT);
    const startedValid = Number.isFinite(started) && started > 0;

    const fmt = function (value) {
      const time = typeof value === 'number' ? value : Date.parse(String(value).replace(/Z$/, '+09:00'));
      return Number.isFinite(time) && time > 0
        ? Utilities.formatDate(new Date(time), SMN.timezone, 'yyyy-MM-dd HH:mm:ss')
        : 'なし';
    };

    console.log('診断時刻（日本時間）: ' + fmt(now));
    console.log('通知開始（日本時間）: ' + fmt(started));
    console.log('通知: ' + (state.SMN_ENABLED === '1' ? '有効' : '停止'));
    console.log('開始時の対象設定: ' +
      (state.SMN_IDENTITY === configIdentity_(c) ? '一致' : '不一致'));
    console.log('対象従業員: ' + (c.staffIds || '全員') +
      ' / 対象事業所: ' + (c.storeIds || 'すべて'));
    console.log('送信待ち: ' + (state[SMN.outboxKey] ? 'あり' : 'なし'));

    const rows = fetchResults_(c, new Date(now));
    const recent = rows.slice().sort(function (a, b) {
      return (Date.parse(b.attendanceAt) || 0) -
        (Date.parse(a.attendanceAt) || 0);
    }).slice(0, 30);

    console.log('勤怠取得: ' + rows.length +
      '件 / 最近の' + recent.length + '件を表示');

    recent.forEach(function (r) {
      let reason = 'APIに退勤時刻なし';
      let recorded = '';
      let relation = '比較不可';

      if (r.leavingAt) {
        try {
          const event = events_([r], c).find(function (e) {
            return e.text.startsWith('【退勤】');
          });

          if (!event) {
            reason = '対象フィルターで除外';
          } else {
            recorded = state[event.key] || '';

            if (startedValid) {
              relation = event.at < started
                ? '通知開始より前' : '通知開始以降';
            }

            if (recorded) {
              reason = '処理済み記録あり（送信・除外の内訳は未保存）';
            } else if (event.at > now) {
              reason = '未来の退勤時刻のため除外';
            } else if (!startedValid) {
              reason = '通知開始時刻の記録が不正';
            } else if (event.at < started) {
              reason = '通知開始より前のため除外';
            } else if (event.at < now - SMN.maxAgeHours * 3600000) {
              reason = '取得対象の時間範囲より古いため除外';
            } else {
              reason = '未処理の通知対象';
            }
          }
        } catch (e) {
          reason = '勤怠検証エラー: ' + clean_(e.message);
        }
      }

      console.log(JSON.stringify({
        氏名: clean_(r.staffName || '従業員ID ' + r.staffId),
        勤務日: clean_(r.shiftDate),
        出勤: fmt(r.attendanceAt),
        API退勤: clean_(r.leavingAt || 'なし'),
        退勤_日本時間: fmt(r.leavingAt),
        退勤_丸め後: fmt(r.leavingAtRounded),
        開始との比較: relation,
        判定: reason,
        処理記録時刻: recorded ? fmt(Number(recorded)) : 'なし'
      }));
    });
  });
}