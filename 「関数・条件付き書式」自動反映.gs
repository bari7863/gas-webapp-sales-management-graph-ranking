function applyFormulasAndConditionalFormatting() {
  // =========================
  // 設定
  // =========================

  // 対象スプレッドシートURLを複数入れる
  const spreadsheetUrls = [
        "https://docs.google.com/spreadsheets/d/1P8nXyk9ruBCDOzmyCGKUst3KMJ6HUOjpL1eqMEssJlM/edit?usp=drivesdk",
        "https://docs.google.com/spreadsheets/d/1miqIx-lrrYnSMl-sNDcTPgrzENgAo0QVKilhWSxsbU0/edit?usp=drivesdk",
        "https://docs.google.com/spreadsheets/d/1giyUwJkvJIzucyQQsvvUdmst9Ons99jcKLL9MsvTE2s/edit?usp=drivesdk",
  ];

  // 対象シート名を複数入れる
  const targetSheetNames = [
    '月次管理シートテンプレ_各人',
    '26年2月'
  ];

  // 関数を入れる列範囲
  const startCol = 8;   // H列
  const endCol   = 38;  // AL列

  // =========================
  // 関数テンプレート
  // {COL} の部分が各列(H, I, J ... AL)に置き換わる
  // =========================
  const formulaMap = {
    4: `=IFERROR(
  IF({COL}3="","",
    LET(
      rawLines, SPLIT({COL}3, CHAR(10), FALSE, TRUE),
      lines, FILTER(rawLines, REGEXMATCH(rawLines, "[0-9].*[\\-ー－〜～~].*[0-9]")),
      hours,
        SUM(
          MAP(
            lines,
            LAMBDA(line,
              LET(
                startText, TRIM(REGEXEXTRACT(line, "^\\s*([^\\-ー－〜～~]+)")),
                endText,   TRIM(REGEXEXTRACT(line, "[\\-ー－〜～~]\\s*(.+)$")),
                startDigits, REGEXREPLACE(startText, "[^0-9]", ""),
                endDigits,   REGEXREPLACE(endText, "[^0-9]", ""),
                startNum, IF(LEN(startDigits)<=2, VALUE(startDigits)*100, VALUE(startDigits)),
                endNum,   IF(LEN(endDigits)<=2, VALUE(endDigits)*100, VALUE(endDigits)),
                (
                  TIME(INT(endNum/100), MOD(endNum,100), 0)
                  -
                  TIME(INT(startNum/100), MOD(startNum,100), 0)
                ) * 24
              )
            )
          )
        ),
      IF(hours >= 6, hours - 1, hours)
    )
  ),
"")`,

    11: `=IF({COL}4="","",INT({COL}4*20))`,

    13: `=IF({COL}11="","",ROUNDDOWN({COL}11*0.07,0))`,

    15: `=IF({COL}11="","",ROUNDDOWN({COL}11*0.03,0))`,

    17: `=IF({COL}11="","",MAX(1,ROUNDDOWN({COL}11/80,0)))`,

    19: `=IFERROR(AVERAGE(FILTER({
  IFERROR({COL}18/{COL}12, NA())
}, ISNUMBER({
  IFERROR({COL}18/{COL}12, NA())
}))), "")`,

    20: `=IFERROR(ROUND(AVERAGE(FILTER({
  IFERROR(({COL}5/{COL}18)*60, NA())
}, ISNUMBER({
  IFERROR(({COL}5/{COL}18)*60, NA())
}))), 0), "")`,

    21: `="お疲れ様です！"&CHAR(10)&CHAR(10)&
"【終業報告】"&
IF(LEN({COL}5),CHAR(10)&"稼働時間："&{COL}5&"時間","")&
IF(LEN({COL}6),CHAR(10)&"商談数："&{COL}6&"件","")&
IF(LEN({COL}12),CHAR(10)&"コール数："&{COL}12&"件","")&
IF(LEN({COL}14),CHAR(10)&"見込み数："&{COL}14&"件","")&
IF(LEN({COL}16),CHAR(10)&"代表接触数："&{COL}16&"件","")&
IF(LEN({COL}18),CHAR(10)&"アポ数："&{COL}18&"件","")&
CHAR(10)&"コール先リスト："`
  };

  // =========================
  // 実行
  // =========================
  spreadsheetUrls.forEach(url => {
    const ss = SpreadsheetApp.openById(getSpreadsheetIdFromUrl_(url));

    targetSheetNames.forEach(sheetName => {
      const sheet = ss.getSheetByName(sheetName);
      if (!sheet) {
        Logger.log(`シートが見つかりません: ${sheetName} / ${url}`);
        return;
      }

      // 関数を設定
      applyFormulasToSheet_(sheet, startCol, endCol, formulaMap);

      // 条件付き書式を設定
      applyConditionalFormattingToSheet_(sheet);

      Logger.log(`完了: ${ss.getName()} / ${sheetName}`);
    });
  });
}


/**
 * 指定シートに関数を入れる
 */
function applyFormulasToSheet_(sheet, startCol, endCol, formulaMap) {
  Object.keys(formulaMap).forEach(rowStr => {
    const row = Number(rowStr);
    const formulas = [];

    for (let col = startCol; col <= endCol; col++) {
      const colLetter = columnToLetter_(col);
      const formula = formulaMap[row].replace(/\{COL\}/g, colLetter);
      formulas.push(formula);
    }

    sheet.getRange(row, startCol, 1, formulas.length).setFormulas([formulas]);
  });
}


/**
 * 指定シートに条件付き書式を設定
 * ※ G11:G12 と G13:G20 の既存ルールは削除して入れ直す
 */
function applyConditionalFormattingToSheet_(sheet) {
  const existingRules = sheet.getConditionalFormatRules();

  const filteredRules = existingRules.filter(rule => {
    const ranges = rule.getRanges();
    const boolCondition = rule.getBooleanCondition();

    const a1Notations = ranges.map(r => r.getA1Notation());

    // ① G11:G12 / G13:G20 の既存ルールは削除
    const isTargetRangeRule = a1Notations.some(a1 =>
      a1 === 'G11:G12' || a1 === 'G13:G20'
    );

    if (isTargetRangeRule) {
      return false;
    }

    // ② G7:G20 の =LEFT(G7,1)="-" ルールがあれば削除
    const isOldMinusRuleRange = a1Notations.some(a1 => a1 === 'G7:G20');

    if (isOldMinusRuleRange && boolCondition) {
      const criteriaType = boolCondition.getCriteriaType();
      const criteriaValues = boolCondition.getCriteriaValues();

      if (
        criteriaType === SpreadsheetApp.BooleanCriteria.CUSTOM_FORMULA &&
        criteriaValues &&
        criteriaValues[0] === '=LEFT(G7,1)="-"'
      ) {
        return false;
      }
    }

    return true;
  });

  const range1 = sheet.getRange('G11:G12');
  const range2 = sheet.getRange('G13:G20');

  const newRules = [
    SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied('=AND(ISNUMBER($G11),$G11<=-100)')
      .setBackground('#ff0000')
      .setRanges([range1])
      .build(),

    SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied('=AND(ISNUMBER($G11),$G11>=-99,$G11<=-50)')
      .setBackground('#ff9900')
      .setRanges([range1])
      .build(),

    SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied('=AND(ISNUMBER($G11),$G11>=-49,$G11<=-1)')
      .setBackground('#ffff00')
      .setRanges([range1])
      .build(),

    SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied('=AND(ISNUMBER($G13),$G13<=-10)')
      .setBackground('#ff0000')
      .setRanges([range2])
      .build(),

    SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied('=AND(ISNUMBER($G13),$G13>=-9,$G13<=-5)')
      .setBackground('#ff9900')
      .setRanges([range2])
      .build(),

    SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied('=AND(ISNUMBER($G13),$G13>=-4,$G13<=-1)')
      .setBackground('#ffff00')
      .setRanges([range2])
      .build()
  ];

  sheet.setConditionalFormatRules([...filteredRules, ...newRules]);
}


/**
 * スプレッドシートURLからIDを抜き出す
 */
function getSpreadsheetIdFromUrl_(url) {
  const match = url.match(/\/d\/([a-zA-Z0-9-_]+)/);
  if (!match) {
    throw new Error(`スプレッドシートURLが不正です: ${url}`);
  }
  return match[1];
}


/**
 * 列番号 → 列記号
 * 例: 8 -> H
 */
function columnToLetter_(column) {
  let temp = '';
  while (column > 0) {
    let remainder = (column - 1) % 26;
    temp = String.fromCharCode(65 + remainder) + temp;
    column = Math.floor((column - 1) / 26);
  }
  return temp;
}
