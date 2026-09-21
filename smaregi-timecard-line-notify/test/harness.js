/**
 * GASの組み込みサービスをNode上で模擬し、pollTimecard の挙動を確認する簡易ハーネス。
 * GASへはデプロイしない。実行例:
 *   node test/harness.js src/コード.js              # 識別方式の移行を含む初回
 *   SEED_V2=1 node test/harness.js src/コード.js    # 移行済みの定常状態
 */
const crypto = require('crypto');
const fs = require('fs');

function makeEnv(rowsRef, sentRef) {
  const store = {};
  const cache = {};
  global.PropertiesService = {
    getScriptProperties: () => ({
      getProperty: k => (Object.prototype.hasOwnProperty.call(store, k) ? store[k] : null),
      setProperty: (k, v) => { store[k] = String(v); },
      getProperties: () => Object.assign({}, store),
      setProperties: (o) => { Object.keys(o).forEach(k => { store[k] = String(o[k]); }); },
      deleteProperty: k => { delete store[k]; }
    })
  };
  global.CacheService = { getScriptCache: () => ({ get: k => cache[k] || null, put: (k, v) => { cache[k] = v; } }) };
  global.LockService = { getScriptLock: () => ({ tryLock: () => true, releaseLock: () => {} }) };
  global.Utilities = {
    DigestAlgorithm: { SHA_256: 'SHA_256' },
    Charset: { UTF_8: 'UTF_8' },
    computeDigest: (_a, s) => Array.from(crypto.createHash('sha256').update(s, 'utf8').digest()).map(b => b > 127 ? b - 256 : b),
    base64Encode: s => Buffer.from(s, 'utf8').toString('base64'),
    getUuid: () => crypto.randomUUID(),
    formatDate: (date, tz, fmt) => {
      const p = {};
      new Intl.DateTimeFormat('en-US', {
        timeZone: tz, year: 'numeric', month: 'numeric', day: 'numeric',
        hour: '2-digit', minute: '2-digit', second: '2-digit', hour12: false
      }).formatToParts(date).forEach(x => { p[x.type] = x.value; });
      const pad = v => String(v).padStart(2, '0');
      if (fmt === 'yyyy-MM-dd') return `${p.year}-${pad(p.month)}-${pad(p.day)}`;
      if (fmt === 'M/d HH:mm:ss') return `${Number(p.month)}/${Number(p.day)} ${pad(p.hour)}:${p.minute}:${p.second}`;
      if (fmt === 'yyyy-MM-dd HH:mm:ss') return `${p.year}-${pad(p.month)}-${pad(p.day)} ${pad(p.hour)}:${p.minute}:${p.second}`;
      throw new Error('unsupported format ' + fmt);
    }
  };
  global.UrlFetchApp = {
    fetch: (url, opts) => {
      const ok = (obj, code = 200, headers = {}) => ({
        getResponseCode: () => code, getContentText: () => JSON.stringify(obj), getAllHeaders: () => headers
      });
      if (url.includes('id.smaregi.jp')) return ok({ access_token: 't0ken', expires_in: 3600 });
      if (url.includes('/timecard/shifts/results')) return ok({ results: rowsRef.rows, page: 1, pageCount: 1 });
      if (url.includes('/message/quota/consumption')) return ok({ totalUsage: 5 });
      if (url.includes('/message/quota')) return ok({ type: 'limited', value: 200 });
      if (url.includes('/message/push')) {
        sentRef.push(JSON.parse(opts.payload).messages[0].text);
        return ok({}, 200, { 'x-line-accepted-request-id': 'r1' });
      }
      if (url.includes('/profile/')) return ok({ displayName: 'owner' });
      throw new Error('unexpected url ' + url);
    }
  };
  return store;
}

function load(file) {
  const src = fs.readFileSync(file, 'utf8');
  // Evaluate in global scope so top-level `const SMN` and functions are reachable.
  (0, eval)(src + '\nglobal.__poll = pollTimecard; global.__events = events_;');
}

function run(label, file, script) {
  const rowsRef = { rows: [] };
  const sent = [];
  const store = makeEnv(rowsRef, sent);
  load(file);
  Object.assign(store, {
    LINE_ACCESS_TOKEN: 'tok', LINE_USER_ID: 'U' + '0'.repeat(32),
    SMAREGI_CONTRACT_ID: 'abc123', SMAREGI_CLIENT_ID: 'cid', SMAREGI_CLIENT_SECRET: 'sec',
    SMN_ENABLED: '1', SMN_STARTED_AT: String(Date.parse('2026-09-21T00:00:00+09:00'))
  });
  if (process.env.SEED_V2 === '1') store.SMN_KEY_VERSION = 'v2';
  store.SMN_IDENTITY = (0, eval)('configIdentity_')((0, eval)('config_')());
  console.log('\n=== ' + label + ' (' + file + ') ===');
  script({ rowsRef, sent, store, poll: () => { sent.length = 0; global.__poll(); return sent.slice(); } });
}

const row = (o) => Object.assign({
  staffId: 7, storeId: 2, shiftDate: '2026-09-21',
  staffName: '山田太郎', storeName: 'LIG高田馬場店', leavingAt: null
}, o);

const NOW = Date.parse('2026-09-21T13:00:00+09:00');
const realNow = Date.now;
Date.now = () => NOW;

for (const file of process.argv.slice(2)) {
  run('打刻修正シナリオ', file, ({ rowsRef, poll }) => {
    rowsRef.rows = [row({ attendanceAt: '2026-09-21T12:47:05Z' })];
    console.log('1回目(出勤 12:47:05):', JSON.stringify(poll()));
    rowsRef.rows = [row({ attendanceAt: '2026-09-21T12:41:00Z' })]; // スマレジ側で打刻修正
    console.log('2回目(12:41:00へ修正):', JSON.stringify(poll()));
    console.log('3回目(変化なし):', JSON.stringify(poll()));
    rowsRef.rows = [row({ attendanceAt: '2026-09-21T12:41:00Z', leavingAt: '2026-09-21T12:55:00Z' })];
    console.log('4回目(退勤追加):', JSON.stringify(poll()));
  });

  run('同日2勤務(中抜け)シナリオ', file, ({ rowsRef, poll }) => {
    rowsRef.rows = [row({ attendanceAt: '2026-09-21T09:00:00Z', leavingAt: '2026-09-21T11:00:00Z' })];
    console.log('1回目(1件目):', JSON.stringify(poll()));
    rowsRef.rows.push(row({ attendanceAt: '2026-09-21T12:41:00Z' }));
    console.log('2回目(2件目の出勤):', JSON.stringify(poll()));
  });
}
Date.now = realNow;
