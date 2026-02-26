// ★メイン：数字だけ更新（グラフはそのまま）
function updateGraphNumbersOnly() {
  const ss  = SpreadsheetApp.getActive();
  const src = ss.getSheetByName('26年2月'); // 必要ならここだけ変更
  if (!src) throw new Error('「26年2月」シートが見つかりません');

  // ★今回のレイアウトに合わせた定義
  const members  = getMembers_(src);         // 1人目=22行目、以降19行間隔（22,41,60...）
  const metrics  = getMetricDefinitions_();  // コール・アポ・受注 + アポ率 + アポ生産性
  const todayCol = getTodayColumn_(src);     // 本日の列（H〜）

  // 「当日グラフ」「月間グラフ」に、数値だけ書き込む
  updateDailySheetValues_(ss, src, members, metrics, todayCol);
  updateMonthlySheetValues_(ss, src, members, metrics);
}

/**
 * 1人目：22行目〜39行目、2人目：41〜58…（1ブロック18行 + 1行空き）
 */
function getMembers_(sheet) {
  var members = [];
  var startRow    = 22; // 1人目の先頭行（A22）
  var blockHeight = 19; // 18行ブロック + 間に1行空き（22→41）
  var lastRow     = sheet.getLastRow();
  var memberCount = Math.floor((lastRow - startRow) / blockHeight) + 1;
  if (memberCount < 0) memberCount = 0;

  for (var i = 0; i < memberCount; i++) {
    var baseRow = startRow + i * blockHeight;
    var name = sheet.getRange(baseRow, 1).getDisplayValue(); // A列
    if (name) {
      members.push({ name: name, baseRow: baseRow });
    }
  }
  return members;
}

/**
 * 指標定義（baseRow を起点にしたオフセット）
 * 0:シフト
 * 2:稼働時間
 * 6,7:必要受注数/受注数
 * 8,9:必要コール数/実コール数
 * 14,15:必要アポ数/アポ数
 * 16:アポ率
 * 17:アポ生産性
 */
function getMetricDefinitions_() {
  // グラフ対象（従来どおり：コール・アポ・受注）
  const call  = { key: 'call',  label: 'コール数', goalOffset: 8,  actualOffset: 9  };
  const apo   = { key: 'apo',   label: 'アポ数',   goalOffset: 14, actualOffset: 15 };
  const order = { key: 'order', label: '受注数',   goalOffset: 6,  actualOffset: 7  };

  // 売上金額（目標: 必要売上金額 / 実数: 売上金額）
  const sales = { key: 'sales', label: '売上金額', goalOffset: 4, actualOffset: 5, numberFormat: '¥#,##0' };

  const pairs = [call, apo, order]; // ※グラフ用（G〜I）はこの3つだけ

  // ペア参照用
  const pairMap = { call, apo, order, sales };

  // ★追加：表（A〜E）の表示順（指定順）
  const tableOrder = ['call', 'apo', 'apoRate', 'apoProductivity', 'order', 'sales'];

  // アポ率（実数のみ）
  const apoRate = {
    key: 'apoRate',
    label: 'アポ率',
    offset: 16
  };

  // アポ生産性（実数のみ）
  const apoProductivity = {
    key: 'apoProductivity',
    label: 'アポ生産性',
    offset: 17,
    numberFormat: '0.0"時""間""/""件"'
  };

  return { pairs, pairMap, tableOrder, apoRate, apoProductivity };
}

/**
 * 文字列でも「0.0時間」「10件」「¥1,000」などから数値を取り出す
 * Date（時間表示のセル）も簡易対応
 */
function normalizeNumber_(val) {
  if (typeof val === 'number') return val;
  if (val === '' || val === null || val === undefined) return 0;

  if (val instanceof Date) {
    return val.getHours() + (val.getMinutes() / 60) + (val.getSeconds() / 3600) + (val.getMilliseconds() / 3600000);
  }

  let s = String(val);
  s = s.replace(/[^\d\.\-]/g, ''); // 数字・小数点・マイナスのみ残す
  if (!s) return 0;

  const n = parseFloat(s);
  return isNaN(n) ? 0 : n;
}

/**
 * 1行目から「今日の日付」の列(H〜)を探す
 *  - 今日と一致する列があればそれを返す
 *  - 見つからなければ「H1以降で一番左にある日付列」を返す
 */
function getTodayColumn_(sheet) {
  const firstCol = 8; // H列
  const lastCol  = sheet.getLastColumn();
  if (lastCol < firstCol) throw new Error('H列以降に日付がありません。');

  const header = sheet.getRange(1, firstCol, 1, lastCol - firstCol + 1).getValues()[0];

  const today = new Date();
  today.setHours(0, 0, 0, 0);

  let fallbackCol = null;

  for (let i = 0; i < header.length; i++) {
    const v = header[i];
    if (v instanceof Date) {
      const d = new Date(v);
      d.setHours(0, 0, 0, 0);

      if (fallbackCol === null) fallbackCol = firstCol + i;
      if (d.getTime() === today.getTime()) return firstCol + i;
    }
  }

  if (fallbackCol !== null) return fallbackCol;
  throw new Error('H1以降に日付がありません。');
}

// 当日グラフ用シートに数字だけ書き込む
// ★グラフはそのまま、A〜E（表）と G〜I（グラフ用データ）だけ更新
function updateDailySheetValues_(ss, src, members, metrics, todayCol) {
  const sh = ss.getSheetByName('当日グラフ');
  if (!sh) throw new Error('「当日グラフ」シートがありません（先にグラフ作成を実行してください）');

  // 数値だけ全消去（グラフは残す）
  sh.getRange(1, 1, sh.getMaxRows(), 5).clearContent(); // A〜E
  sh.getRange(1, 7, sh.getMaxRows(), 3).clearContent(); // G〜I

  let row = 1;

  // ───────────────
  // ① A〜E：表（コール・アポ・受注・アポ率・アポ生産性）
  // ───────────────
  const tableData = [['氏名', '項目', '', '目標', '実数']];

  members.forEach(m => {
    metrics.tableOrder.forEach(k => {
      if (k === 'apoRate') {
        const apoRateVal = normalizeNumber_(src.getRange(m.baseRow + metrics.apoRate.offset, todayCol).getValue());
        tableData.push([m.name, metrics.apoRate.label, m.name + ' / ' + metrics.apoRate.label, '', apoRateVal]);
        return;
      }

      if (k === 'apoProductivity') {
        const apoProdVal = normalizeNumber_(src.getRange(m.baseRow + metrics.apoProductivity.offset, todayCol).getValue());
        tableData.push([m.name, metrics.apoProductivity.label, m.name + ' / ' + metrics.apoProductivity.label, '', apoProdVal]);
        return;
      }

      const metric = metrics.pairMap[k];
      const goal = normalizeNumber_(src.getRange(m.baseRow + metric.goalOffset,   todayCol).getValue());
      const act  = normalizeNumber_(src.getRange(m.baseRow + metric.actualOffset, todayCol).getValue());
      tableData.push([m.name, metric.label, m.name + ' / ' + metric.label, goal, act]);
    });
  });

  sh.getRange(row, 1, tableData.length, tableData[0].length).setValues(tableData);

  // フォーマット（1人あたり6行：コール・アポ・アポ率・アポ生産性・受注・売上金額）
  for (let i = 0; i < members.length; i++) {
    const blockStartRow = row + 1 + i * 6;

    // コール・アポ（2行）D/E を整数
    sh.getRange(blockStartRow, 4, 2, 2).setNumberFormat('0');

    // アポ率（3行目）E を %
    sh.getRange(blockStartRow + 2, 5, 1, 1).setNumberFormat('0.0%');

    // アポ生産性（4行目）E を指定フォーマット
    sh.getRange(blockStartRow + 3, 5, 1, 1).setNumberFormat(metrics.apoProductivity.numberFormat);

    // 受注数（5行目）D/E を整数
    sh.getRange(blockStartRow + 4, 4, 1, 2).setNumberFormat('0');

    // 売上金額（6行目）D/E を通貨フォーマット
    sh.getRange(blockStartRow + 5, 4, 1, 2).setNumberFormat(metrics.pairMap.sales.numberFormat);
  }

  // ───────────────
  // ② G〜I：グラフ用（コール・アポ・受注のみ）
  // ───────────────
  const chartData = [['', '目標', '実数']];

  members.forEach(m => {
    metrics.pairs.forEach(metric => {
      const goal = normalizeNumber_(src.getRange(m.baseRow + metric.goalOffset,   todayCol).getValue());
      const act  = normalizeNumber_(src.getRange(m.baseRow + metric.actualOffset, todayCol).getValue());
      chartData.push([m.name + ' / ' + metric.label, goal, act]);
    });
  });

  sh.getRange(row, 7, chartData.length, chartData[0].length).setValues(chartData);
  if (chartData.length > 1) {
    sh.getRange(row + 1, 8, chartData.length - 1, 2).setNumberFormat('0'); // H/I を整数
  }

  // ───────────────
  // ③ 1枚目チャート：参照範囲とタイトルだけ更新（設定は維持）
  // ───────────────
  const charts = sh.getCharts();
  if (charts.length >= 1) {
    const tz = ss.getSpreadsheetTimeZone();
    const dateStr = Utilities.formatDate(new Date(), tz, 'yyyy/MM/dd');

    const b = charts[0].modify();
    b.clearRanges();
    b.addRange(sh.getRange(row, 7, chartData.length, 3)); // G〜I
    b.setOption('title', '当日（' + dateStr + '） コール・アポ・受注 目標 vs 実数');
    sh.updateChart(b.build());
  }
}

// 月間グラフ用シートに数字だけ書き込む
// ★グラフはそのまま、A〜E（表）と G〜I（グラフ用データ）だけ更新
function updateMonthlySheetValues_(ss, src, members, metrics) {
  const sh = ss.getSheetByName('月間グラフ');
  if (!sh) throw new Error('「月間グラフ」シートがありません（先にグラフ作成を実行してください）');

  // 数値だけ全消去（グラフは残す）
  sh.getRange(1, 1, sh.getMaxRows(), 5).clearContent();
  sh.getRange(1, 7, sh.getMaxRows(), 3).clearContent();

  const monthlyCol = 6; // F列＝月間合計
  let row = 1;

  // ───────────────
  // ① A〜E：表（コール・アポ・受注・アポ率・アポ生産性）
  // ───────────────
  const tableData = [['氏名', '項目', '', '目標', '実数']];

  members.forEach(m => {
    metrics.tableOrder.forEach(k => {
      if (k === 'apoRate') {
        const apoRateVal = normalizeNumber_(src.getRange(m.baseRow + metrics.apoRate.offset, monthlyCol).getValue());
        tableData.push([m.name, metrics.apoRate.label, m.name + ' / ' + metrics.apoRate.label, '', apoRateVal]);
        return;
      }

      if (k === 'apoProductivity') {
        const apoProdVal = normalizeNumber_(src.getRange(m.baseRow + metrics.apoProductivity.offset, monthlyCol).getValue());
        tableData.push([m.name, metrics.apoProductivity.label, m.name + ' / ' + metrics.apoProductivity.label, '', apoProdVal]);
        return;
      }

      const metric = metrics.pairMap[k];
      const goal = normalizeNumber_(src.getRange(m.baseRow + metric.goalOffset,   monthlyCol).getValue());
      const act  = normalizeNumber_(src.getRange(m.baseRow + metric.actualOffset, monthlyCol).getValue());
      tableData.push([m.name, metric.label, m.name + ' / ' + metric.label, goal, act]);
    });
  });

  sh.getRange(row, 1, tableData.length, tableData[0].length).setValues(tableData);

  // フォーマット（1人あたり6行）
  for (let i = 0; i < members.length; i++) {
    const blockStartRow = row + 1 + i * 6;

    // コール・アポ（2行）D/E を整数
    sh.getRange(blockStartRow, 4, 2, 2).setNumberFormat('0');

    // アポ率（3行目）E を %
    sh.getRange(blockStartRow + 2, 5, 1, 1).setNumberFormat('0.0%');

    // アポ生産性（4行目）E を指定フォーマット
    sh.getRange(blockStartRow + 3, 5, 1, 1).setNumberFormat(metrics.apoProductivity.numberFormat);

    // 受注数（5行目）D/E を整数
    sh.getRange(blockStartRow + 4, 4, 1, 2).setNumberFormat('0');

    // 売上金額（6行目）D/E を通貨フォーマット
    sh.getRange(blockStartRow + 5, 4, 1, 2).setNumberFormat(metrics.pairMap.sales.numberFormat);
  }

  // ───────────────
  // ② G〜I：グラフ用（コール・アポ・受注のみ）
  // ───────────────
  const chartData = [['', '目標', '実数']];

  members.forEach(m => {
    metrics.pairs.forEach(metric => {
      const goal = normalizeNumber_(src.getRange(m.baseRow + metric.goalOffset,   monthlyCol).getValue());
      const act  = normalizeNumber_(src.getRange(m.baseRow + metric.actualOffset, monthlyCol).getValue());
      chartData.push([m.name + ' / ' + metric.label, goal, act]);
    });
  });

  sh.getRange(row, 7, chartData.length, chartData[0].length).setValues(chartData);
  if (chartData.length > 1) {
    sh.getRange(row + 1, 8, chartData.length - 1, 2).setNumberFormat('0');
  }

  // ───────────────
  // ③ 1枚目チャート：参照範囲とタイトルだけ更新
  // ───────────────
  const charts = sh.getCharts();
  if (charts.length >= 1) {
    const b = charts[0].modify();
    b.clearRanges();
    b.addRange(sh.getRange(row, 7, chartData.length, 3));
    b.setOption('title', '月間累計 コール・アポ・受注 目標 vs 実数');
    sh.updateChart(b.build());
  }
}
