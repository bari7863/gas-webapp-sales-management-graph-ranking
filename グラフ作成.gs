/**
 * メイン関数：
 * ・当日分 → シート「当日グラフ」
 * ・月間累計 → シート「月間グラフ」
 * に「コール数・アポ数・受注数」の縦棒グラフを作成
 * ※アポ率・アポ生産性は表にだけ追加（グラフ対象外）
 */
function createDailyAndMonthlyCharts() {
  var ss  = SpreadsheetApp.getActiveSpreadsheet();
  var src = ss.getSheetByName('26年2月');
  if (!src) throw new Error('シート「26年2月」が見つかりません。');

  var members       = getMembers_(src);
  var metrics       = getMetricDefinitions_();
  var spreadsheetTz = ss.getSpreadsheetTimeZone();

  // 当日用
  createDailyGraphSheet_(ss, src, members, metrics, spreadsheetTz);

  // 月間用
  createMonthlyGraphSheet_(ss, src, members, metrics);
}

function getMembers_(sheet) {
  var members = [];
  // 1人目：22〜39、2人目：41〜58 … 15人目：288〜305
  // → baseRow は 22, 41, 60... と 19行間隔
  var startRow    = 22; // 1人目の先頭行（A22）
  var blockHeight = 19; // 18行ブロック + 間に1行空き（22→41）
  var memberCount = 15;

  for (var i = 0; i < memberCount; i++) {
    var baseRow = startRow + i * blockHeight;
    var name = sheet.getRange(baseRow, 1).getDisplayValue(); // A列
    if (name) {
      members.push({ name: name, baseRow: baseRow });
    }
  }
  return members;
}

function getMetricDefinitions_() {
  // 稼働時間（実績のみ）
  var worktime = {
    key:   'worktime',
    label: '稼働時間',
    offset: 2
  };

  // グラフ対象（3つ）
  var call  = { key: 'call',  label: 'コール数', goalOffset: 8,  actualOffset: 9  };
  var apo   = { key: 'apo',   label: 'アポ数',   goalOffset: 14, actualOffset: 15 };
  var order = { key: 'order', label: '受注数',   goalOffset: 6,  actualOffset: 7  };

  // 売上金額（目標: 必要売上金額 / 実数: 売上金額）
  var sales = { key: 'sales', label: '売上金額', goalOffset: 4, actualOffset: 5, numberFormat: '¥#,##0' };

  // グラフ対象（従来どおり：コール数・アポ数・受注数）
  var pairs = [call, apo, order];

  // ペアメトリクス参照用
  var pairMap = { call: call, apo: apo, order: order, sales: sales };

  // 表の表示順（指定順）
  var tableOrder = ['call', 'apo', 'apoRate', 'apoProductivity', 'order', 'sales'];

  // アポ率（目標なし・実数のみ）
  var apoRate = {
    key:    'apoRate',
    label:  'アポ率',
    offset: 16
  };

  // アポ生産性（目標なし・実数のみ）
  // アポ率の次の行を想定 → baseRow + 17
  var apoProductivity = {
    key:    'apoProductivity',
    label:  'アポ生産性',
    offset: 17,
    numberFormat: '0.0"時""間""/""件"'
  };

  return {
    worktime: worktime,
    pairs: pairs,
    pairMap: pairMap,
    tableOrder: tableOrder,
    apoRate: apoRate,
    apoProductivity: apoProductivity
  };
}

/**
 * 文字列でも「0.0時間」「10件」「¥1,000」などから数値を取り出す
 * Date（時間表示のセル）も一応対応
 */
function normalizeNumber_(val) {
  if (typeof val === 'number') return val;
  if (val === '' || val === null || val === undefined) return 0;

  // Duration/Time が Date で返るケースの簡易対応（時間として扱う）
  if (val instanceof Date) {
    return val.getHours() + (val.getMinutes() / 60) + (val.getSeconds() / 3600) + (val.getMilliseconds() / 3600000);
  }

  var s = String(val);
  s = s.replace(/[^\d\.\-]/g, ''); // 数字・小数点・マイナスのみ残す
  if (!s) return 0;

  var n = parseFloat(s);
  if (isNaN(n)) return 0;
  return n;
}

/**
 * 1行目から「今日の日付」の列(H〜)を探す
 *  - 今日と一致する列があればそれを返す
 *  - 見つからなければ「H1以降で一番左にある日付列」を返す
 */
function getTodayColumn_(sheet) {
  var firstCol = 8; // H列スタート
  var lastCol  = sheet.getLastColumn();
  if (lastCol < firstCol) throw new Error('H列以降に日付がありません。');

  var headerRowValues = sheet.getRange(1, firstCol, 1, lastCol - firstCol + 1).getValues()[0];

  var today = new Date();
  today.setHours(0, 0, 0, 0);

  var fallbackCol = null;

  for (var i = 0; i < headerRowValues.length; i++) {
    var v = headerRowValues[i];

    if (v instanceof Date) {
      var d = new Date(v);
      d.setHours(0, 0, 0, 0);

      if (fallbackCol === null) {
        fallbackCol = firstCol + i;
      }

      if (d.getTime() === today.getTime()) {
        return firstCol + i;
      }
    }
  }

  if (fallbackCol !== null) return fallbackCol;

  throw new Error('H1以降に日付がありません。');
}

/**
 * シートを取得（あれば初期化、なければ作成）
 */
function prepareSheet_(ss, sheetName) {
  var sheet = ss.getSheetByName(sheetName);
  if (!sheet) {
    sheet = ss.insertSheet(sheetName);
  } else {
    sheet.clear();
    var charts = sheet.getCharts();
    charts.forEach(function(c) { sheet.removeChart(c); });
  }
  return sheet;
}

/**
 * 当日分：
 * ・A〜E列：コール数・アポ数・受注数・アポ率・アポ生産性（表）
 * ・G〜I列：コール数・アポ数・受注数のみ（グラフ用データ）
 * ※アポ率・アポ生産性はグラフ対象外
 */
function createDailyGraphSheet_(ss, src, members, metrics, spreadsheetTz) {
  var dest     = prepareSheet_(ss, '当日グラフ');
  var todayCol = getTodayColumn_(src);
  var dateStr  = Utilities.formatDate(new Date(), spreadsheetTz, 'yyyy/MM/dd');

  var row = 1;

  // 表（A〜E）
  var tableData = [['氏名', '項目', '', '目標', '実数']];

  // tableOrder / pairMap が未定義でも動くようにフォールバックを用意
  var tableOrder = metrics.tableOrder || ['call', 'apo', 'apoRate', 'apoProductivity', 'order', 'sales'];

  var pairMap = metrics.pairMap || (function() {
    var m = {};
    (metrics.pairs || []).forEach(function(p) { m[p.key] = p; });
    return m;
  })();

  // 売上金額が pairMap に無い場合も定義しておく
  if (!pairMap.sales) {
    pairMap.sales = { key: 'sales', label: '売上金額', goalOffset: 4, actualOffset: 5, numberFormat: '¥#,##0' };
  }

  members.forEach(function(m) {
    tableOrder.forEach(function(k) {
      if (k === 'apoRate') {
        var rawApo = src.getRange(m.baseRow + metrics.apoRate.offset, todayCol).getValue();
        var apoVal = normalizeNumber_(rawApo);
        tableData.push([m.name, metrics.apoRate.label, m.name + ' / ' + metrics.apoRate.label, '', apoVal]);
        return;
      }

      if (k === 'apoProductivity') {
        var rawApoProd = src.getRange(m.baseRow + metrics.apoProductivity.offset, todayCol).getValue();
        var apoProdVal = normalizeNumber_(rawApoProd);
        tableData.push([m.name, metrics.apoProductivity.label, m.name + ' / ' + metrics.apoProductivity.label, '', apoProdVal]);
        return;
      }

      var metric  = pairMap[k];
      var rawGoal = src.getRange(m.baseRow + metric.goalOffset,   todayCol).getValue();
      var rawAct  = src.getRange(m.baseRow + metric.actualOffset, todayCol).getValue();
      var goal    = normalizeNumber_(rawGoal);
      var act     = normalizeNumber_(rawAct);

      tableData.push([m.name, metric.label, m.name + ' / ' + metric.label, goal, act]);
    });
  });

  var tableRows = tableData.length;
  dest.getRange(row, 1, tableRows, tableData[0].length).setValues(tableData);

  // 数値フォーマット（1人あたり6行：コール・アポ・アポ率・アポ生産性・受注・売上金額）
  for (var i = 0; i < members.length; i++) {
    var blockStartRow = row + 1 + i * 6; // ヘッダーの次が1人目の先頭

    // コール・アポ（2行） D・E を整数
    dest.getRange(blockStartRow, 4, 2, 2).setNumberFormat('0');

    // アポ率（3行目）E を 0.0%
    var apoRow = blockStartRow + 2;
    dest.getRange(apoRow, 5, 1, 1).setNumberFormat('0.0%');

    // アポ生産性（4行目）E を指定フォーマット
    var apoProdRow = blockStartRow + 3;
    dest.getRange(apoProdRow, 5, 1, 1).setNumberFormat(metrics.apoProductivity.numberFormat);

    // 受注数（5行目）D・E を整数
    var orderRow = blockStartRow + 4;
    dest.getRange(orderRow, 4, 1, 2).setNumberFormat('0');

    // 売上金額（6行目）D・E を通貨フォーマット
    var salesRow = blockStartRow + 5;
    dest.getRange(salesRow, 4, 1, 2).setNumberFormat(pairMap.sales.numberFormat);
  }

  // グラフ用データ（G〜I）
  var chartData = [['', '目標', '実数']];

  members.forEach(function(m) {
    metrics.pairs.forEach(function(metric) {
      var rawGoal = src.getRange(m.baseRow + metric.goalOffset,   todayCol).getValue();
      var rawAct  = src.getRange(m.baseRow + metric.actualOffset, todayCol).getValue();
      var goal    = normalizeNumber_(rawGoal);
      var act     = normalizeNumber_(rawAct);

      chartData.push([m.name + ' / ' + metric.label, goal, act]);
    });
  });

  var chartRows = chartData.length;
  dest.getRange(row, 7, chartRows, chartData[0].length).setValues(chartData);
  if (chartRows > 1) {
    dest.getRange(row + 1, 8, chartRows - 1, 2).setNumberFormat('0'); // H・I
  }

  // 当日 縦棒グラフ（G〜Iのみ）
  var chart = dest.newChart()
    .setChartType(Charts.ChartType.COLUMN)
    .addRange(dest.getRange(row, 7, chartRows, 3))
    .setOption('title', '当日（' + dateStr + '） コール・アポ・受注 目標 vs 実数')
    .setOption('legend', {
      position: 'top',
      textStyle: { fontSize: 14, bold: true }
    })
    .setOption('series', {
      0: { color: '#ff6b6b' }, // 目標
      1: { color: '#4f8dfb' }  // 実数
    })
    .setOption('annotations', {
      alwaysOutside: false,
      textStyle: { color: '#ffffff', bold: true, fontSize: 12 }
    })
    .setOption('hAxis', { textStyle: { fontSize: 12, bold: false } })
    .setOption('vAxis', {
      textStyle: { fontSize: 12, bold: false },
      logScale: true,
      format: '0'
    })
    .setOption('titleTextStyle', { bold: true, fontSize: 14 })
    .setOption('bar', { groupWidth: '80%' })
    .setOption('width', 1200)
    .setOption('height', 800)
    .setPosition(row, 11, 0, 0)
    .build();

  dest.insertChart(chart);
}

/**
 * 月間：
 * ・A〜E列：コール数・アポ数・受注数・アポ率・アポ生産性（表）
 * ・G〜I列：コール数・アポ数・受注数のみ（グラフ用データ）
 * ※アポ率・アポ生産性はグラフ対象外
 */
function createMonthlyGraphSheet_(ss, src, members, metrics) {
  var dest = prepareSheet_(ss, '月間グラフ');
  var row = 1;

  // 表（A〜E）
  var tableData = [['氏名', '項目', '', '目標', '実数']];

  // tableOrder / pairMap が未定義でも動くようにフォールバックを用意
  var tableOrder = metrics.tableOrder || ['call', 'apo', 'apoRate', 'apoProductivity', 'order', 'sales'];

  var pairMap = metrics.pairMap || (function() {
    var m = {};
    (metrics.pairs || []).forEach(function(p) { m[p.key] = p; });
    return m;
  })();

  // 売上金額が pairMap に無い場合も定義しておく
  if (!pairMap.sales) {
    pairMap.sales = { key: 'sales', label: '売上金額', goalOffset: 4, actualOffset: 5, numberFormat: '¥#,##0' };
  }

  members.forEach(function(m) {
    tableOrder.forEach(function(k) {
      if (k === 'apoRate') {
        var rawApo = src.getRange(m.baseRow + metrics.apoRate.offset, 6).getValue();
        var apoVal = normalizeNumber_(rawApo);
        tableData.push([m.name, metrics.apoRate.label, m.name + ' / ' + metrics.apoRate.label, '', apoVal]);
        return;
      }

      if (k === 'apoProductivity') {
        var rawApoProd = src.getRange(m.baseRow + metrics.apoProductivity.offset, 6).getValue();
        var apoProdVal = normalizeNumber_(rawApoProd);
        tableData.push([m.name, metrics.apoProductivity.label, m.name + ' / ' + metrics.apoProductivity.label, '', apoProdVal]);
        return;
      }

      var metric  = pairMap[k];
      var rawGoal = src.getRange(m.baseRow + metric.goalOffset,   6).getValue(); // F列
      var rawAct  = src.getRange(m.baseRow + metric.actualOffset, 6).getValue();
      var goal    = normalizeNumber_(rawGoal);
      var act     = normalizeNumber_(rawAct);

      tableData.push([m.name, metric.label, m.name + ' / ' + metric.label, goal, act]);
    });
  });

  var tableRows = tableData.length;
  dest.getRange(row, 1, tableRows, tableData[0].length).setValues(tableData);

  // 数値フォーマット（1人あたり6行）
  for (var i = 0; i < members.length; i++) {
    var blockStartRow = row + 1 + i * 6;

    // コール・アポ（2行） D・E は整数
    dest.getRange(blockStartRow, 4, 2, 2).setNumberFormat('0');

    // アポ率（3行目）E は 0.0%
    var apoRow = blockStartRow + 2;
    dest.getRange(apoRow, 5, 1, 1).setNumberFormat('0.0%');

    // アポ生産性（4行目）E は指定フォーマット
    var apoProdRow = blockStartRow + 3;
    dest.getRange(apoProdRow, 5, 1, 1).setNumberFormat(metrics.apoProductivity.numberFormat);

    // 受注数（5行目）D・E は整数
    var orderRow = blockStartRow + 4;
    dest.getRange(orderRow, 4, 1, 2).setNumberFormat('0');

    // 売上金額（6行目）D・E は通貨フォーマット
    var salesRow = blockStartRow + 5;
    dest.getRange(salesRow, 4, 1, 2).setNumberFormat(pairMap.sales.numberFormat);
  }

  // グラフ用データ（G〜I）
  var chartData = [['', '目標', '実数']];

  members.forEach(function(m) {
    metrics.pairs.forEach(function(metric) {
      var rawGoal = src.getRange(m.baseRow + metric.goalOffset,   6).getValue();
      var rawAct  = src.getRange(m.baseRow + metric.actualOffset, 6).getValue();
      var goal    = normalizeNumber_(rawGoal);
      var act     = normalizeNumber_(rawAct);

      chartData.push([m.name + ' / ' + metric.label, goal, act]);
    });
  });

  var chartRows = chartData.length;
  dest.getRange(row, 7, chartRows, chartData[0].length).setValues(chartData);
  if (chartRows > 1) {
    dest.getRange(row + 1, 8, chartRows - 1, 2).setNumberFormat('0');
  }

  // 月間 縦棒グラフ（G〜Iのみ）
  var chart = dest.newChart()
    .setChartType(Charts.ChartType.COLUMN)
    .addRange(dest.getRange(row, 7, chartRows, 3))
    .setOption('title', '月間累計 コール・アポ・受注 目標 vs 実数')
    .setOption('legend', {
      position: 'top',
      textStyle: { fontSize: 14, bold: true }
    })
    .setOption('series', {
      0: { color: '#ff6b6b' }, // 目標
      1: { color: '#4f8dfb' }  // 実数
    })
    .setOption('annotations', {
      alwaysOutside: false,
      textStyle: { color: '#ffffff', bold: true, fontSize: 12 }
    })
    .setOption('hAxis', { textStyle: { fontSize: 12, bold: false } })
    .setOption('vAxis', {
      textStyle: { fontSize: 12, bold: false },
      logScale: true,
      format: '0'
    })
    .setOption('titleTextStyle', { bold: true, fontSize: 14 })
    .setOption('bar', { groupWidth: '80%' })
    .setOption('width', 1200)
    .setOption('height', 800)
    .setPosition(row, 11, 0, 0)
    .build();

  dest.insertChart(chart);
}

/* ここから下は（今回の変更点とは無関係）残してOK */

function writeDailySingleMetric_(dest, src, members, metric, todayCol, startRow) {
  var headerRow    = startRow;
  var dataStartRow = headerRow + 1;

  dest.getRange(headerRow, 1, 1, 2).setValues([
    ['氏名', metric.label + ' 実数']
  ]);

  var values = [];
  members.forEach(function(m) {
    var raw = src.getRange(m.baseRow + metric.offset, todayCol).getValue();
    var val = normalizeNumber_(raw);
    values.push([m.name, val]);
  });

  dest.getRange(dataStartRow, 1, values.length, 2).setValues(values);

  var chart = dest.newChart()
    .setChartType(Charts.ChartType.COLUMN)
    .addRange(dest.getRange(headerRow, 1, values.length + 1, 2))
    .setOption('title', '【当日】' + metric.label + '（実数）')
    .setOption('legend', { position: 'none' })
    .setOption('width', 900)
    .setOption('height', 400)
    .setPosition(headerRow, 5, 0, 0)
    .build();

  dest.insertChart(chart);

  return dataStartRow + values.length + 25;
}

function writeDailyMetricPair_(dest, src, members, metric, todayCol, startRow) {
  var headerRow    = startRow;
  var dataStartRow = headerRow + 1;

  dest.getRange(headerRow, 1, 1, 3).setValues([[
    '氏名',
    metric.label + ' 目標',
    metric.label + ' 実数'
  ]]);

  var values = [];
  members.forEach(function(m) {
    var rawGoal = src.getRange(m.baseRow + metric.goalOffset,   todayCol).getValue();
    var rawAct  = src.getRange(m.baseRow + metric.actualOffset, todayCol).getValue();
    var goal    = normalizeNumber_(rawGoal);
    var act     = normalizeNumber_(rawAct);
    values.push([m.name, goal, act]);
  });

  dest.getRange(dataStartRow, 1, values.length, 3).setValues(values);

  var chart = dest.newChart()
    .setChartType(Charts.ChartType.COLUMN)
    .addRange(dest.getRange(headerRow, 1, values.length + 1, 3))
    .setOption('title', '【当日】' + metric.label + '（目標 vs 実数）')
    .setOption('legend', { position: 'bottom' })
    .setOption('series', {
      0: { color: '#0000ff' },
      1: { color: '#ff0000' }
    })
    .setOption('width', 900)
    .setOption('height', 400)
    .setPosition(headerRow, 5, 0, 0)
    .build();

  dest.insertChart(chart);

  return dataStartRow + values.length + 25;
}

function writeMonthlySingleMetric_(dest, src, members, metric, startRow) {
  var headerRow    = startRow;
  var dataStartRow = headerRow + 1;

  dest.getRange(headerRow, 1, 1, 2).setValues([
    ['氏名', metric.label + '（月間 実数）']
  ]);

  var values = [];
  members.forEach(function(m) {
    var raw = src.getRange(m.baseRow + metric.offset, 6).getValue();
    var val = normalizeNumber_(raw);
    values.push([m.name, val]);
  });

  dest.getRange(dataStartRow, 1, values.length, 2).setValues(values);

  var chart = dest.newChart()
    .setChartType(Charts.ChartType.COLUMN)
    .addRange(dest.getRange(headerRow, 1, values.length + 1, 2))
    .setOption('title', '【月間】' + metric.label + '（実数）')
    .setOption('legend', { position: 'none' })
    .setOption('width', 900)
    .setOption('height', 400)
    .setPosition(headerRow, 5, 0, 0)
    .build();

  dest.insertChart(chart);

  return dataStartRow + values.length + 25;
}

function writeMonthlyMetricPair_(dest, src, members, metric, startRow) {
  var headerRow    = startRow;
  var dataStartRow = headerRow + 1;

  dest.getRange(headerRow, 1, 1, 3).setValues([[
    '氏名',
    metric.label + ' 目標（月間）',
    metric.label + ' 実数（月間）'
  ]]);

  var values = [];
  members.forEach(function(m) {
    var rawGoal = src.getRange(m.baseRow + metric.goalOffset,   6).getValue();
    var rawAct  = src.getRange(m.baseRow + metric.actualOffset, 6).getValue();
    var goal    = normalizeNumber_(rawGoal);
    var act     = normalizeNumber_(rawAct);
    values.push([m.name, goal, act]);
  });

  dest.getRange(dataStartRow, 1, values.length, 3).setValues(values);

  var chart = dest.newChart()
    .setChartType(Charts.ChartType.COLUMN)
    .addRange(dest.getRange(headerRow, 1, values.length + 1, 3))
    .setOption('title', '【月間】' + metric.label + '（目標 vs 実数）')
    .setOption('legend', { position: 'bottom' })
    .setOption('series', {
      0: { color: '#0000ff' },
      1: { color: '#ff0000' }
    })
    .setOption('width', 900)
    .setOption('height', 400)
    .setPosition(headerRow, 5, 0, 0)
    .build();

  dest.insertChart(chart);

  return dataStartRow + values.length + 25;
}
