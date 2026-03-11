/***** 設定（ここだけ書き換え） *****/
const LOGTEST_CONF = {
  LARK: {
    APP_ID: '',
    APP_SECRET: '',
    APP_TOKEN: '', // Bitable App token
    TABLE_ID: ''             // Table ID
  },
  SHEET: {
    SPREADSHEET_ID: '',
    SHEET_NAME: '商談ログ管理',
  }
};

// 営業マンはこのリストのみ
const LOGTEST_SALESMEN = [
  //退職

  //現職
  'ココちゃん',
  'バリ道場',
  'バリ食堂',
];

/***** 初回だけ実行（A列プルダウン + 条件付き書式で色分け）*****/
function setupLogTestSalesmanDropdownAndColors() {
  const ss = SpreadsheetApp.openById(LOGTEST_CONF.SHEET.SPREADSHEET_ID);
  const sh = ss.getSheetByName(LOGTEST_CONF.SHEET.SHEET_NAME);
  if (!sh) throw new Error('シートが見つかりません: ' + LOGTEST_CONF.SHEET.SHEET_NAME);

  const maxRows = sh.getMaxRows();

  // A2:A にプルダウン制限
  const rangeA = sh.getRange(2, 1, Math.max(maxRows - 1, 1), 1); // A2:A
  const rule = SpreadsheetApp.newDataValidation()
    .requireValueInList(LOGTEST_SALESMEN, true)
    .setAllowInvalid(false)
    .build();
  rangeA.setDataValidation(rule);

  // 色（好きに変更OK）
  const colors = [
    '#FFF2CC', '#D9EAD3', '#D0E0E3', '#CFE2F3', '#D9D2E9',
    '#FCE5CD', '#F4CCCC', '#EAD1DC', '#D9E1F2', '#E2EFDA',
    '#FFF2F2', '#EAF4FF', '#F3E5F5', '#E8F5E9'
  ];

  // 既存の条件付き書式から「この営業マン色分けルールっぽいもの」だけ除外して作り直す
  const existing = sh.getConditionalFormatRules();
  const kept = existing.filter(r => {
    const bc = r.getBooleanCondition();
    if (!bc) return true;
    if (bc.getCriteriaType() !== SpreadsheetApp.BooleanCriteria.TEXT_EQUAL_TO) return true;
    const v = bc.getCriteriaValues()[0];
    if (!LOGTEST_SALESMEN.includes(v)) return true;

    // 対象範囲にA列が含まれてるものだけを「色分け候補」とみなす
    return !r.getRanges().some(rr => rr.getColumn() === 1);
  });

  const newRules = LOGTEST_SALESMEN.map((name, i) => {
    return SpreadsheetApp.newConditionalFormatRule()
      .whenTextEqualTo(name)
      .setBackground(colors[i % colors.length])
      .setRanges([rangeA]) // A列だけ色分け。行全体にしたいならA2:Cに変更
      .build();
  });

  sh.setConditionalFormatRules([...kept, ...newRules]);
}

/***** 手動実行用（この関数を実行してください） *****/
function runLarkToSheet_ShodanLogTest() {
  LOGTEST_syncOverwrite_();
}

/***** 上書き同期（A:営業マン / B:商談日 / C:案件名） *****/
function LOGTEST_syncOverwrite_() {
  const ss = SpreadsheetApp.openById(LOGTEST_CONF.SHEET.SPREADSHEET_ID);
  const sh = ss.getSheetByName(LOGTEST_CONF.SHEET.SHEET_NAME);
  if (!sh) throw new Error('シートが見つかりません: ' + LOGTEST_CONF.SHEET.SHEET_NAME);

  // ヘッダー
  sh.getRange(1, 1, 1, 3).setValues([['営業マン', '商談日', '案件名']]);

  // ★ ここが抜けていたのがエラー原因：Larkからレコード取得
  const items = LOGTEST_fetchAllLarkRecords_();

  // 行データ作成（営業マンがリスト外ならスキップ）
  const allowSet = new Set(LOGTEST_SALESMEN);

  const rows = items
    .map(it => {
      const f = it.fields || {};
      const salesman = LOGTEST_normalize_(f['営業マン']);
      if (!allowSet.has(salesman)) return null;

      const dt = LOGTEST_toDateOnly_(f['商談日時（訪問日時）']); // 日付だけ
      const company = LOGTEST_normalize_(f['企業名']);
      return [salesman, dt, company];
    })
    .filter(Boolean);

  // 既存データをクリア（A:C 2行目以降）
  const startRow = 2;
  const last = sh.getLastRow();
  if (last >= startRow) {
    sh.getRange(startRow, 1, last - startRow + 1, 3).clearContent();
  }

  // 書き込み
  if (rows.length) {
    sh.getRange(startRow, 1, rows.length, 3).setValues(rows);

    // B列を日付表示（時間なし）
    sh.getRange(startRow, 2, rows.length, 1).setNumberFormat('yyyy/mm/dd');
  }

  Logger.log('同期完了: %s件', rows.length);
}

/* ================== Lark 認証/取得 ================== */

/** Lark tenant_access_token 取得 */
function LOGTEST_getTenantAccessToken_() {
  const url = 'https://open.larksuite.com/open-apis/auth/v3/tenant_access_token/internal';
  const res = UrlFetchApp.fetch(url, {
    method: 'post',
    contentType: 'application/json',
    payload: JSON.stringify({
      app_id: LOGTEST_CONF.LARK.APP_ID,
      app_secret: LOGTEST_CONF.LARK.APP_SECRET
    }),
    muteHttpExceptions: true
  });

  const text = res.getContentText();
  const json = JSON.parse(text || '{}');
  if (!json.tenant_access_token) {
    throw new Error('Lark token取得失敗: ' + text);
  }
  return json.tenant_access_token;
}

/** Lark records 全件取得（500件ずつ） */
function LOGTEST_fetchAllLarkRecords_() {
  const token = LOGTEST_getTenantAccessToken_();
  const base = `https://open.larksuite.com/open-apis/bitable/v1/apps/${LOGTEST_CONF.LARK.APP_TOKEN}/tables/${LOGTEST_CONF.LARK.TABLE_ID}/records`;

  let pageToken = null;
  let items = [];

  while (true) {
    const qs = pageToken
      ? `page_token=${encodeURIComponent(pageToken)}&page_size=500&field_key=field_name`
      : `page_size=500&field_key=field_name`;

    const url = `${base}?${qs}`;

    const res = UrlFetchApp.fetch(url, {
      method: 'get',
      headers: {
        Authorization: 'Bearer ' + token,
        'X-Field-Key': 'field_name'
      },
      muteHttpExceptions: true
    });

    const text = res.getContentText();
    const json = JSON.parse(text || '{}');

    if (!json.data || !json.data.items) {
      throw new Error('Lark records取得失敗: ' + text);
    }

    items = items.concat(json.data.items);

    if (!json.data.has_more || !json.data.page_token) break;
    pageToken = json.data.page_token;
  }

  return items;
}

/* ================== 値の整形 ================== */

function LOGTEST_normalize_(v) {
  if (v == null) return '';
  if (v instanceof Date) return v;

  const t = typeof v;
  if (t === 'string' || t === 'number' || t === 'boolean') return String(v).trim();

  if (Array.isArray(v)) {
    return v.map(LOGTEST_normalize_).filter(x => x !== '').join(', ');
  }

  if (t === 'object') {
    if (v.text != null) return String(v.text).trim();
    if (v.name != null) return String(v.name).trim();
    if (v.value != null) return LOGTEST_normalize_(v.value);

    if (Array.isArray(v.users)) {
      return v.users.map(u => u.name || u.email || u.id).filter(Boolean).join(', ');
    }

    return JSON.stringify(v);
  }

  return String(v);
}

/** 商談日時（訪問日時）を「日付だけ（時間なし）」の Date にする */
function LOGTEST_toDateOnly_(v) {
  if (v == null || v === '') return '';

  let d = null;

  if (v instanceof Date) {
    d = v;
  } else if (typeof v === 'number') {
    d = new Date(v);
  } else {
    const s = LOGTEST_normalize_(v).replace(/-/g, '/');
    const tmp = new Date(s);
    if (!isNaN(tmp.getTime())) d = tmp;
  }

  if (!d || isNaN(d.getTime())) return LOGTEST_normalize_(v);

  // 時間を落として「その日の 00:00」
  return new Date(d.getFullYear(), d.getMonth(), d.getDate());
}
