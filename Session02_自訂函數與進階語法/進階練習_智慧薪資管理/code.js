// ============================================================
// 進階練習：智慧薪資計算與加班費自動化
// 主題：辦公室自動化 AI — 薪資管理
// 對應：Session 2（自訂函數、除錯、if/for）
// ============================================================

/**
 * 計算加班費（自訂函數）
 * 勞基法規定：
 *   前 2 小時：時薪 × 1.34
 *   第 3~4 小時：時薪 × 1.67
 *   假日加班前 8 小時：時薪 × 1.34
 *   假日加班超過 8 小時：時薪 × 1.67
 *
 * @param {number} 月薪 - 員工月薪
 * @param {number} 加班時數 - 總加班時數
 * @param {string} 加班類型 - "平日" 或 "假日"
 * @returns {number} 加班費
 * @customfunction
 */
function 計算加班費(月薪, 加班時數, 加班類型) {
  if (typeof 月薪 !== "number" || typeof 加班時數 !== "number") {
    return "❌ 請輸入數字";
  }
  if (加班時數 < 0) return "❌ 時數不能為負";

  // 時薪 = 月薪 / 30 / 8
  var 時薪 = 月薪 / 30 / 8;
  var 加班費 = 0;

  if (加班類型 === "假日") {
    // 假日加班
    if (加班時數 <= 8) {
      加班費 = 加班時數 * 時薪 * 1.34;
    } else {
      加班費 = 8 * 時薪 * 1.34 + (加班時數 - 8) * 時薪 * 1.67;
    }
  } else {
    // 平日加班
    if (加班時數 <= 2) {
      加班費 = 加班時數 * 時薪 * 1.34;
    } else {
      加班費 = 2 * 時薪 * 1.34 + (加班時數 - 2) * 時薪 * 1.67;
    }
  }

  return Math.round(加班費);
}

/**
 * 計算勞健保自付額（自訂函數）
 * @param {number} 月薪 - 員工月薪
 * @returns {number} 勞保 + 健保自付額
 * @customfunction
 */
function 計算勞健保(月薪) {
  if (typeof 月薪 !== "number") return "❌ 請輸入數字";

  // 簡化版：勞保費率約 11.5% 員工自付 20%，健保費率 5.17% 員工自付 30%
  var 投保薪資 = _取得投保級距(月薪);
  var 勞保自付 = Math.round(投保薪資 * 0.115 * 0.2);
  var 健保自付 = Math.round(投保薪資 * 0.0517 * 0.3);

  return 勞保自付 + 健保自付;
}

/**
 * 取得勞保投保級距（簡化版）
 * @param {number} 月薪
 * @returns {number} 對應的投保薪資
 */
function _取得投保級距(月薪) {
  var 級距 = [27470, 28800, 30300, 31800, 33300, 34800, 36300,
              38200, 40100, 42000, 43900, 45800];
  for (var i = 0; i < 級距.length; i++) {
    if (月薪 <= 級距[i]) return 級距[i];
  }
  return 45800; // 最高級距
}

/**
 * 計算所得稅預扣（自訂函數）
 * @param {number} 月薪 - 月薪
 * @param {number} 扶養人數 - 扶養親屬人數
 * @returns {number} 預扣所得稅
 * @customfunction
 */
function 計算所得稅(月薪, 扶養人數) {
  if (typeof 月薪 !== "number") return "❌ 請輸入數字";
  if (!扶養人數) 扶養人數 = 0;

  // 簡化版免稅額計算
  var 免稅額 = 92000 + (扶養人數 * 92000);
  var 年收 = 月薪 * 12;
  var 標準扣除 = 124000;
  var 薪資特扣 = 207000;

  var 應稅所得 = Math.max(0, 年收 - 免稅額 - 標準扣除 - 薪資特扣);

  // 累進稅率
  var 年稅;
  if (應稅所得 <= 560000) {
    年稅 = 應稅所得 * 0.05;
  } else if (應稅所得 <= 1260000) {
    年稅 = 560000 * 0.05 + (應稅所得 - 560000) * 0.12;
  } else if (應稅所得 <= 2520000) {
    年稅 = 560000 * 0.05 + 700000 * 0.12 + (應稅所得 - 1260000) * 0.2;
  } else {
    年稅 = 560000 * 0.05 + 700000 * 0.12 + 1260000 * 0.2 + (應稅所得 - 2520000) * 0.3;
  }

  return Math.round(年稅 / 12); // 月預扣
}

/**
 * 計算實發薪資（自訂函數）
 * @param {number} 月薪
 * @param {number} 加班費
 * @param {number} 扶養人數
 * @returns {number} 實發金額
 * @customfunction
 */
function 實發薪資(月薪, 加班費, 扶養人數) {
  if (typeof 月薪 !== "number") return "❌ 請輸入數字";
  if (!加班費) 加班費 = 0;
  if (!扶養人數) 扶養人數 = 0;

  var 勞健保 = 計算勞健保(月薪);
  var 所得稅 = 計算所得稅(月薪, 扶養人數);

  return 月薪 + 加班費 - 勞健保 - 所得稅;
}

// ============================================================
// 批次計算全公司薪資
// ============================================================

/**
 * 一鍵計算全部員工薪資
 */
function 批次計算薪資() {
  try {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("薪資明細");
    if (!sheet) {
      SpreadsheetApp.getUi().alert("❌ 請先執行「初始化薪資資料」");
      return;
    }

    var 資料 = sheet.getDataRange().getValues();
    var 結果 = [];

    for (var i = 1; i < 資料.length; i++) {
      var 月薪 = 資料[i][2];       // C: 月薪
      var 平日加班 = 資料[i][3];   // D: 平日加班時數
      var 假日加班 = 資料[i][4];   // E: 假日加班時數
      var 扶養 = 資料[i][5];       // F: 扶養人數

      var 平日加班費 = 計算加班費(月薪, 平日加班, "平日");
      var 假日加班費 = 計算加班費(月薪, 假日加班, "假日");
      var 總加班費 = 平日加班費 + 假日加班費;
      var 勞健保 = 計算勞健保(月薪);
      var 所得稅 = 計算所得稅(月薪, 扶養);
      var 實發 = 月薪 + 總加班費 - 勞健保 - 所得稅;

      結果.push([平日加班費, 假日加班費, 總加班費, 勞健保, 所得稅, 實發]);
    }

    // 批次寫入
    if (結果.length > 0) {
      sheet.getRange(2, 7, 結果.length, 6).setValues(結果);
      sheet.getRange(2, 7, 結果.length, 6).setNumberFormat("#,##0");
    }

    // 加入合計列
    var 合計列 = 資料.length + 1;
    sheet.getRange(合計列, 1).setValue("📊 合計").setFontWeight("bold");
    for (var col = 7; col <= 12; col++) {
      var 欄字 = String.fromCharCode(64 + col);
      sheet.getRange(合計列, col).setFormula(
        "=SUM(" + 欄字 + "2:" + 欄字 + (合計列 - 1) + ")"
      );
    }
    sheet.getRange(合計列, 1, 1, 12).setBackground("#e8f5e9").setFontWeight("bold");
    sheet.getRange(合計列, 7, 1, 6).setNumberFormat("#,##0");

    SpreadsheetApp.getUi().alert("✅ 全公司薪資計算完成！\n共 " + 結果.length + " 位員工");

  } catch (錯誤) {
    Logger.log("❌ 錯誤：" + 錯誤.message);
    SpreadsheetApp.getUi().alert("❌ 錯誤：" + 錯誤.message);
  }
}

// ============================================================
// 初始化與選單
// ============================================================

function 初始化薪資資料() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("薪資明細");
  if (!sheet) sheet = ss.insertSheet("薪資明細"); else sheet.clear();

  var 標題 = [["姓名", "部門", "月薪", "平日加班(hr)", "假日加班(hr)", "扶養人數",
               "平日加班費", "假日加班費", "加班費合計", "勞健保", "所得稅", "實發薪資"]];
  var 資料 = [
    ["王小明", "業務部", 42000, 12, 8, 0, "", "", "", "", "", ""],
    ["李小華", "行銷部", 38000, 8, 0, 2, "", "", "", "", "", ""],
    ["張美玲", "人資部", 45000, 6, 4, 1, "", "", "", "", "", ""],
    ["陳大文", "研發部", 58000, 20, 16, 0, "", "", "", "", "", ""],
    ["林小芬", "財務部", 40000, 4, 0, 3, "", "", "", "", "", ""],
    ["黃志偉", "業務部", 52000, 15, 12, 1, "", "", "", "", "", ""],
    ["劉家豪", "研發部", 65000, 25, 8, 0, "", "", "", "", "", ""],
    ["吳雅琪", "行銷部", 36000, 10, 4, 0, "", "", "", "", "", ""],
    ["周建國", "業務部", 35000, 8, 8, 2, "", "", "", "", "", ""],
    ["許文馨", "人資部", 48000, 6, 0, 1, "", "", "", "", "", ""]
  ];

  sheet.getRange(1, 1, 1, 12).setValues(標題);
  sheet.getRange(2, 1, 資料.length, 12).setValues(資料);
  sheet.getRange("A1:L1").setBackground("#0d47a1").setFontColor("#fff").setFontWeight("bold")
    .setHorizontalAlignment("center");
  sheet.getRange("C2:C11").setNumberFormat("#,##0");
  sheet.setFrozenRows(1);
  for (var c = 1; c <= 12; c++) sheet.autoResizeColumn(c);

  SpreadsheetApp.getUi().alert("✅ 薪資資料已建立！\n請執行「批次計算薪資」計算結果。");
}

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu("🤖 智慧薪資管理")
    .addItem("📦 初始化薪資資料", "初始化薪資資料")
    .addItem("🧮 批次計算薪資", "批次計算薪資")
    .addToUi();
}
