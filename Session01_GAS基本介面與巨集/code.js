// ============================================================
// Session 1：Google Apps Script 基本介面與巨集
// 日期：115/05/02（六）09:00~12:00
// 講師：林冠廷
// ============================================================
// 本課程涵蓋：
//   1. GAS 基本介面認識
//   2. 基本語法（變數、函數、流程控制）
//   3. 巨集 (Macro) 錄製與編輯
//   4. 觸發器 (Trigger) 初步應用
// ============================================================

// ============================================================
// 第一部分：基本語法 — 變數與資料型別
// ============================================================

/**
 * 示範基本變數宣告與資料型別
 * 學習重點：var、let、const 的差異，以及常見資料型別
 */
function 基本變數示範() {
  // --- 變數宣告方式 ---

  // var：函數作用域（舊式寫法，不建議使用）
  var 姓名 = "王小明";

  // let：區塊作用域（可重新賦值）
  let 年齡 = 25;

  // const：區塊作用域（不可重新賦值，建議常數使用）
  const 公司名稱 = "ABC科技公司";

  // --- 資料型別 ---
  let 文字 = "Hello World";      // 字串 (String)
  let 數字 = 42;                  // 數字 (Number)
  let 小數 = 3.14;                // 數字 (Number)
  let 布林值 = true;              // 布林 (Boolean)
  let 空值 = null;                // 空值 (Null)
  let 未定義;                     // 未定義 (Undefined)

  // 使用 Logger.log 輸出結果（類似 console.log）
  Logger.log("姓名：" + 姓名);
  Logger.log("年齡：" + 年齡);
  Logger.log("公司：" + 公司名稱);
  Logger.log("文字型別：" + typeof 文字);
  Logger.log("數字型別：" + typeof 數字);
  Logger.log("布林型別：" + typeof 布林值);
}

// ============================================================
// 第二部分：函數 (Function)
// ============================================================

/**
 * 基本函數宣告與呼叫
 * 學習重點：函數定義、參數傳遞、回傳值
 */
function 函數示範() {
  // 呼叫自定義函數
  let 結果 = 計算加總(10, 20);
  Logger.log("10 + 20 = " + 結果);

  // 呼叫打招呼函數
  let 問候語 = 打招呼("同學");
  Logger.log(問候語);
}

/**
 * 計算兩數加總
 * @param {number} a - 第一個數字
 * @param {number} b - 第二個數字
 * @returns {number} 兩數之和
 */
function 計算加總(a, b) {
  return a + b;
}

/**
 * 產生問候語
 * @param {string} 名稱 - 要問候的對象
 * @returns {string} 問候語句
 */
function 打招呼(名稱) {
  return "嗨！" + 名稱 + "，歡迎來到 Google Apps Script 課程！";
}

// ============================================================
// 第三部分：流程控制 — if / else / switch
// ============================================================

/**
 * 示範條件判斷語法
 * 學習重點：if-else、比較運算子、邏輯運算子
 */
function 流程控制示範() {
  // 從試算表讀取成績資料來做判斷
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("成績單");

  if (!sheet) {
    Logger.log("❌ 找不到「成績單」工作表，請先建立範例資料");
    return;
  }

  // 取得所有資料（從第 2 列開始，跳過標題）
  var 資料範圍 = sheet.getDataRange();
  var 所有資料 = 資料範圍.getValues();

  Logger.log("===== 成績評等結果 =====");

  // 從第 1 列開始（索引 1），跳過標題列（索引 0）
  for (var i = 1; i < 所有資料.length; i++) {
    var 學生姓名 = 所有資料[i][0];
    var 國文 = 所有資料[i][1];
    var 英文 = 所有資料[i][2];
    var 數學 = 所有資料[i][3];
    var 平均 = (國文 + 英文 + 數學) / 3;

    // if-else 判斷等第
    var 等第;
    if (平均 >= 90) {
      等第 = "優等 ⭐";
    } else if (平均 >= 80) {
      等第 = "甲等";
    } else if (平均 >= 70) {
      等第 = "乙等";
    } else if (平均 >= 60) {
      等第 = "丙等";
    } else {
      等第 = "不及格 ⚠️";
    }

    Logger.log(學生姓名 + " — 平均：" + 平均.toFixed(1) + " — 等第：" + 等第);
  }
}

// ============================================================
// 第四部分：巨集 (Macro) — 手動建立的巨集範例
// ============================================================

/**
 * 巨集範例：一鍵格式化標題列
 * 說明：這個函數模擬「錄製巨集」後所產生的程式碼
 *       學生可以先用「工具 > 巨集 > 錄製巨集」自動產生，
 *       再手動修改以學習程式碼結構
 */
function 格式化標題列巨集() {
  try {
    var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = spreadsheet.getSheetByName("成績單");

    if (!sheet) {
      SpreadsheetApp.getUi().alert("❌ 找不到「成績單」工作表！");
      return;
    }

    // 選取標題列（第 1 列）
    var 標題範圍 = sheet.getRange("A1:F1");

    // 設定背景顏色（深藍色）
    標題範圍.setBackground("#1a73e8");

    // 設定文字顏色（白色）
    標題範圍.setFontColor("#ffffff");

    // 設定字體大小
    標題範圍.setFontSize(12);

    // 設定粗體
    標題範圍.setFontWeight("bold");

    // 設定置中對齊
    標題範圍.setHorizontalAlignment("center");

    // 設定欄位自動調整寬度
    for (var col = 1; col <= 6; col++) {
      sheet.autoResizeColumn(col);
    }

    // 凍結標題列
    sheet.setFrozenRows(1);

    Logger.log("✅ 標題列格式化完成！");
    SpreadsheetApp.getUi().alert("✅ 標題列格式化完成！");

  } catch (錯誤) {
    Logger.log("❌ 錯誤：" + 錯誤.message);
    SpreadsheetApp.getUi().alert("❌ 發生錯誤：" + 錯誤.message);
  }
}

/**
 * 巨集範例：一鍵加入框線
 * 說明：為所有資料範圍加上框線
 */
function 加入框線巨集() {
  try {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("成績單");

    if (!sheet) {
      SpreadsheetApp.getUi().alert("❌ 找不到「成績單」工作表！");
      return;
    }

    // 取得含有資料的範圍
    var 資料範圍 = sheet.getDataRange();

    // 設定所有框線（上下左右 + 內部）
    資料範圍.setBorder(true, true, true, true, true, true,
      "#000000", SpreadsheetApp.BorderStyle.SOLID);

    Logger.log("✅ 框線已加入！");
    SpreadsheetApp.getUi().alert("✅ 框線已加入！");

  } catch (錯誤) {
    Logger.log("❌ 錯誤：" + 錯誤.message);
  }
}

// ============================================================
// 第五部分：觸發器 (Trigger) 初步應用
// ============================================================

/**
 * onOpen 觸發器：試算表開啟時自動執行
 * 說明：這是一個「簡單觸發器」，會在使用者開啟試算表時
 *       自動建立自訂選單
 */
function onOpen() {
  var ui = SpreadsheetApp.getUi();

  // 建立自訂選單「📚 課程工具」
  ui.createMenu("📚 課程工具")
    .addItem("🎨 格式化標題列", "格式化標題列巨集")
    .addItem("📏 加入框線", "加入框線巨集")
    .addSeparator()  // 分隔線
    .addItem("📊 計算成績等第", "流程控制示範")
    .addItem("📝 產生成績報告", "產生成績報告")
    .addToUi();

  Logger.log("✅ 自訂選單已建立");
}

/**
 * 產生成績報告（搭配觸發器使用的進階範例）
 * 說明：讀取成績單資料，在新工作表中產生摘要報告
 */
function 產生成績報告() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var 成績表 = ss.getSheetByName("成績單");

    if (!成績表) {
      SpreadsheetApp.getUi().alert("❌ 找不到「成績單」工作表！");
      return;
    }

    // 檢查是否已存在報告工作表，若存在則刪除重建
    var 舊報告 = ss.getSheetByName("成績報告");
    if (舊報告) {
      ss.deleteSheet(舊報告);
    }

    // 建立新的「成績報告」工作表
    var 報告表 = ss.insertSheet("成績報告");

    // 取得成績資料
    var 資料 = 成績表.getDataRange().getValues();

    // 寫入報告標題
    報告表.getRange("A1").setValue("📊 成績摘要報告");
    報告表.getRange("A1").setFontSize(16).setFontWeight("bold");
    報告表.getRange("A2").setValue("產生時間：" + new Date().toLocaleString("zh-TW"));

    // 計算統計資料
    var 總人數 = 資料.length - 1; // 扣掉標題列
    var 國文總分 = 0, 英文總分 = 0, 數學總分 = 0;
    var 及格人數 = 0;

    for (var i = 1; i < 資料.length; i++) {
      國文總分 += 資料[i][1];
      英文總分 += 資料[i][2];
      數學總分 += 資料[i][3];

      var 平均 = (資料[i][1] + 資料[i][2] + 資料[i][3]) / 3;
      if (平均 >= 60) 及格人數++;
    }

    // 寫入統計表格
    報告表.getRange("A4:B4").setValues([["統計項目", "數值"]]);
    報告表.getRange("A4:B4").setBackground("#1a73e8").setFontColor("#fff").setFontWeight("bold");

    var 統計資料 = [
      ["總人數", 總人數],
      ["國文平均", (國文總分 / 總人數).toFixed(1)],
      ["英文平均", (英文總分 / 總人數).toFixed(1)],
      ["數學平均", (數學總分 / 總人數).toFixed(1)],
      ["及格人數", 及格人數],
      ["及格率", (及格人數 / 總人數 * 100).toFixed(1) + "%"]
    ];

    報告表.getRange(5, 1, 統計資料.length, 2).setValues(統計資料);

    // 自動調整欄寬
    報告表.autoResizeColumn(1);
    報告表.autoResizeColumn(2);

    Logger.log("✅ 成績報告已產生！");
    SpreadsheetApp.getUi().alert("✅ 成績報告已產生！請查看「成績報告」工作表。");

  } catch (錯誤) {
    Logger.log("❌ 錯誤：" + 錯誤.message);
    SpreadsheetApp.getUi().alert("❌ 發生錯誤：" + 錯誤.message);
  }
}

// ============================================================
// 🔧 實作練習：建立範例資料（供教學使用）
// ============================================================

/**
 * 初始化範例資料
 * 說明：如果「成績單」工作表不存在，自動建立並填入範例資料
 */
function 初始化範例資料() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("成績單");

  // 如果工作表不存在，建立新的
  if (!sheet) {
    sheet = ss.insertSheet("成績單");
  } else {
    sheet.clear(); // 清除既有資料
  }

  // 標題列
  var 標題 = [["姓名", "國文", "英文", "數學", "自然", "社會"]];

  // 範例資料
  var 資料 = [
    ["王小明", 85, 92, 78, 88, 90],
    ["李小華", 72, 68, 95, 82, 75],
    ["張美玲", 90, 88, 82, 91, 87],
    ["陳大文", 65, 55, 60, 58, 62],
    ["林小芬", 95, 97, 93, 96, 98],
    ["黃志偉", 78, 82, 70, 75, 80],
    ["劉家豪", 58, 45, 52, 60, 55],
    ["吳雅琪", 88, 85, 90, 87, 92],
    ["周建國", 73, 70, 68, 72, 74],
    ["許文馨", 92, 90, 88, 94, 91]
  ];

  // 寫入標題
  sheet.getRange(1, 1, 1, 6).setValues(標題);

  // 寫入資料
  sheet.getRange(2, 1, 資料.length, 6).setValues(資料);

  Logger.log("✅ 範例資料已建立！共 " + 資料.length + " 筆資料");
  SpreadsheetApp.getUi().alert("✅ 範例資料已建立！共 " + 資料.length + " 筆學生資料");
}
