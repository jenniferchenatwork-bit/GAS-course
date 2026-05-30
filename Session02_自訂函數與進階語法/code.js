// ============================================================
// Session 2：自訂函數 (Custom Functions) 與進階語法
// 日期：115/05/02（六）13:30~16:30
// 講師：林冠廷
// ============================================================
// 本課程涵蓋：
//   1. 自訂函數 (Custom Functions) 建立與應用
//   2. Apps Script 執行限制與除錯
//   3. 進階語法 (if、for)
//   4. 實作：撰寫計算折扣／稅率函數
// ============================================================

// ============================================================
// 第一部分：自訂函數 (Custom Functions)
// 說明：自訂函數可以像 =SUM() 一樣直接在儲存格中使用
//       例如在儲存格輸入 =計算折扣(1000, 0.8)
// ============================================================

/**
 * 計算折扣後金額
 * 使用方式：在 Google Sheet 儲存格輸入 =計算折扣(原價, 折數)
 * 範例：=計算折扣(1000, 0.8) → 回傳 800
 *
 * @param {number} 原價 - 商品原始價格
 * @param {number} 折數 - 折扣比例（0~1 之間，例如 0.8 代表 8 折）
 * @returns {number} 折扣後金額
 * @customfunction
 */
function 計算折扣(原價, 折數) {
  // 參數檢查：確保輸入有效
  if (typeof 原價 !== "number" || typeof 折數 !== "number") {
    return "❌ 請輸入數字";
  }

  if (折數 < 0 || 折數 > 1) {
    return "❌ 折數請輸入 0~1 之間";
  }

  // 計算折扣後的金額
  var 折扣金額 = 原價 * 折數;

  // 回傳結果（四捨五入到整數）
  return Math.round(折扣金額);
}

/**
 * 計算含稅金額（加上營業稅 5%）
 * 使用方式：=計算含稅(1000) → 回傳 1050
 *
 * @param {number} 未稅金額 - 未含稅的金額
 * @param {number} [稅率=0.05] - 稅率（預設 5%）
 * @returns {number} 含稅金額
 * @customfunction
 */
function 計算含稅(未稅金額, 稅率) {
  // 如果沒有傳入稅率，使用預設值 5%
  if (稅率 === undefined || 稅率 === null || 稅率 === "") {
    稅率 = 0.05;
  }

  if (typeof 未稅金額 !== "number") {
    return "❌ 請輸入數字";
  }

  // 計算含稅金額
  var 含稅金額 = 未稅金額 * (1 + 稅率);

  return Math.round(含稅金額);
}

/**
 * 計算折扣後含稅金額（組合函數）
 * 使用方式：=最終金額(1000, 0.8) → 先打 8 折再加 5% 稅
 *
 * @param {number} 原價 - 商品原始價格
 * @param {number} 折數 - 折扣比例
 * @param {number} [稅率=0.05] - 稅率（預設 5%）
 * @returns {number} 最終金額
 * @customfunction
 */
function 最終金額(原價, 折數, 稅率) {
  if (稅率 === undefined || 稅率 === null || 稅率 === "") {
    稅率 = 0.05;
  }

  // 先計算折扣
  var 折扣後 = 計算折扣(原價, 折數);

  // 再加上稅金
  var 含稅 = 計算含稅(折扣後, 稅率);

  return 含稅;
}

/**
 * 判斷商品利潤等級
 * 使用方式：=利潤等級(成本, 售價)
 *
 * @param {number} 成本 - 商品成本
 * @param {number} 售價 - 商品售價
 * @returns {string} 利潤等級描述
 * @customfunction
 */
function 利潤等級(成本, 售價) {
  if (typeof 成本 !== "number" || typeof 售價 !== "number") {
    return "❌ 請輸入數字";
  }

  var 利潤率 = (售價 - 成本) / 成本 * 100;

  // 使用 if-else 判斷利潤等級
  if (利潤率 >= 50) {
    return "🟢 高利潤 (" + 利潤率.toFixed(1) + "%)";
  } else if (利潤率 >= 20) {
    return "🟡 中利潤 (" + 利潤率.toFixed(1) + "%)";
  } else if (利潤率 >= 0) {
    return "🟠 低利潤 (" + 利潤率.toFixed(1) + "%)";
  } else {
    return "🔴 虧損 (" + 利潤率.toFixed(1) + "%)";
  }
}

// ============================================================
// 第二部分：進階語法 — for 迴圈
// ============================================================

/**
 * 使用 for 迴圈批次處理訂單資料
 * 說明：讀取「訂單明細」工作表，計算每筆訂單的折扣、稅金與最終金額
 */
function 批次計算訂單() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName("訂單明細");

    if (!sheet) {
      SpreadsheetApp.getUi().alert("❌ 找不到「訂單明細」工作表！\n請先執行「初始化訂單資料」");
      return;
    }

    // 取得所有資料
    var 資料 = sheet.getDataRange().getValues();
    var 結果 = [];

    Logger.log("===== 開始批次計算 =====");

    // for 迴圈：從第 1 列開始（跳過標題列 index 0）
    for (var i = 1; i < 資料.length; i++) {
      var 商品名 = 資料[i][0];   // A 欄：商品名稱
      var 單價 = 資料[i][1];     // B 欄：單價
      var 數量 = 資料[i][2];     // C 欄：數量
      var 折數 = 資料[i][3];     // D 欄：折扣

      // 計算小計、折扣金額、稅金、合計
      var 小計 = 單價 * 數量;
      var 折扣後 = Math.round(小計 * 折數);
      var 稅金 = Math.round(折扣後 * 0.05);
      var 合計 = 折扣後 + 稅金;

      // 將結果存入陣列
      結果.push([小計, 折扣後, 稅金, 合計]);

      Logger.log(商品名 + "：小計=" + 小計 + " 折後=" + 折扣後 +
                 " 稅金=" + 稅金 + " 合計=" + 合計);
    }

    // 一次性寫入所有結果（效能最佳化）
    if (結果.length > 0) {
      sheet.getRange(2, 5, 結果.length, 4).setValues(結果);
    }

    // 計算訂單總金額
    var 訂單總額 = 0;
    for (var j = 0; j < 結果.length; j++) {
      訂單總額 += 結果[j][3]; // 合計欄
    }

    // 在最後一列寫入總計
    var 總計列 = 資料.length + 1;
    sheet.getRange(總計列, 1).setValue("📊 總計");
    sheet.getRange(總計列, 1).setFontWeight("bold");
    sheet.getRange(總計列, 8).setValue(訂單總額);
    sheet.getRange(總計列, 8).setFontWeight("bold").setBackground("#e8f5e9");

    Logger.log("✅ 批次計算完成！訂單總額：" + 訂單總額);
    SpreadsheetApp.getUi().alert("✅ 批次計算完成！\n訂單總額：NT$ " + 訂單總額.toLocaleString());

  } catch (錯誤) {
    Logger.log("❌ 錯誤：" + 錯誤.message);
    SpreadsheetApp.getUi().alert("❌ 發生錯誤：" + 錯誤.message);
  }
}

// ============================================================
// 第三部分：Apps Script 執行限制與除錯
// ============================================================

/**
 * 示範除錯技巧
 * 說明：展示常見錯誤與除錯方法
 */
function 除錯示範() {
  Logger.log("===== 除錯技巧示範 =====");

  // 技巧 1：使用 Logger.log 追蹤變數值
  var 數值 = 100;
  Logger.log("目前數值：" + 數值);  // 查看變數在某個時間點的值

  // 技巧 2：使用 typeof 檢查資料型別
  var 測試值 = "123";  // 這是字串，不是數字！
  Logger.log("測試值的型別：" + typeof 測試值);  // "string"
  Logger.log("轉換為數字：" + Number(測試值));    // 123

  // 技巧 3：使用 try-catch 捕捉錯誤
  try {
    // 故意製造錯誤：存取不存在的工作表
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("不存在的表");
    var data = sheet.getDataRange().getValues();  // 這裡會出錯！
  } catch (error) {
    Logger.log("捕捉到錯誤：" + error.message);
    Logger.log("錯誤類型：" + error.name);
    // 在實際程式中，可以在這裡做錯誤處理
  }

  // 技巧 4：使用 console.time 測量執行時間
  var 開始時間 = new Date().getTime();

  // 模擬一些處理
  for (var i = 0; i < 10000; i++) {
    var temp = Math.sqrt(i);
  }

  var 結束時間 = new Date().getTime();
  Logger.log("迴圈執行時間：" + (結束時間 - 開始時間) + " 毫秒");

  // ⚠️ Apps Script 執行限制提醒
  Logger.log("");
  Logger.log("===== ⚠️ Apps Script 執行限制 =====");
  Logger.log("• 腳本最大執行時間：6 分鐘");
  Logger.log("• 自訂函數最大執行時間：30 秒");
  Logger.log("• 每日觸發器執行時間：90 分鐘");
  Logger.log("• 每次 UrlFetch 呼叫限制：50 MB");
  Logger.log("• 每日 Email 發送上限：100 封 (免費帳號)");
}

// ============================================================
// 第四部分：進階迴圈與條件判斷組合
// ============================================================

/**
 * 進階範例：巢狀迴圈 + 條件判斷
 * 說明：分析訂單資料，產生不同價位區間的統計
 */
function 訂單分析() {
  try {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("訂單明細");
    if (!sheet) {
      SpreadsheetApp.getUi().alert("❌ 找不到「訂單明細」工作表！");
      return;
    }

    var 資料 = sheet.getDataRange().getValues();

    // 定義價位區間
    var 統計 = {
      "低價 (< $500)": { 數量: 0, 金額: 0 },
      "中價 ($500~$2000)": { 數量: 0, 金額: 0 },
      "高價 (> $2000)": { 數量: 0, 金額: 0 }
    };

    // 遍歷每筆訂單
    for (var i = 1; i < 資料.length; i++) {
      var 單價 = 資料[i][1];
      var 數量 = 資料[i][2];
      var 小計 = 單價 * 數量;

      // 條件判斷分類
      if (小計 < 500) {
        統計["低價 (< $500)"].數量++;
        統計["低價 (< $500)"].金額 += 小計;
      } else if (小計 <= 2000) {
        統計["中價 ($500~$2000)"].數量++;
        統計["中價 ($500~$2000)"].金額 += 小計;
      } else {
        統計["高價 (> $2000)"].數量++;
        統計["高價 (> $2000)"].金額 += 小計;
      }
    }

    // 輸出統計結果
    Logger.log("===== 📊 訂單價位分析 =====");
    for (var 區間 in 統計) {
      Logger.log(區間 + "：" + 統計[區間].數量 + " 筆，總金額 NT$" + 統計[區間].金額);
    }

  } catch (錯誤) {
    Logger.log("❌ 錯誤：" + 錯誤.message);
  }
}

// ============================================================
// 🔧 實作：初始化訂單資料
// ============================================================

/**
 * 建立「訂單明細」工作表與範例資料
 */
function 初始化訂單資料() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("訂單明細");

  if (!sheet) {
    sheet = ss.insertSheet("訂單明細");
  } else {
    sheet.clear();
  }

  // 標題列
  var 標題 = [["商品名稱", "單價", "數量", "折數", "小計", "折扣後", "稅金", "合計"]];
  sheet.getRange(1, 1, 1, 8).setValues(標題);

  // 格式化標題
  var 標題範圍 = sheet.getRange("A1:H1");
  標題範圍.setBackground("#1a73e8");
  標題範圍.setFontColor("#ffffff");
  標題範圍.setFontWeight("bold");
  標題範圍.setHorizontalAlignment("center");

  // 範例資料（只填入前 4 欄，後 4 欄由程式計算）
  var 資料 = [
    ["A4 影印紙 (5包)", 450, 3, 0.9, "", "", "", ""],
    ["原子筆 (盒)", 120, 10, 0.85, "", "", "", ""],
    ["L型資料夾 (50入)", 380, 2, 0.95, "", "", "", ""],
    ["標籤貼紙", 65, 20, 1.0, "", "", "", ""],
    ["釘書機", 180, 5, 0.9, "", "", "", ""],
    ["膠帶台", 250, 3, 0.88, "", "", "", ""],
    ["白板筆 (組)", 320, 4, 0.8, "", "", "", ""],
    ["計算機", 890, 2, 0.75, "", "", "", ""],
    ["筆記本 (10入)", 280, 6, 0.9, "", "", "", ""],
    ["迴紋針 (盒)", 35, 15, 1.0, "", "", "", ""]
  ];

  sheet.getRange(2, 1, 資料.length, 8).setValues(資料);

  // 自動調整欄寬
  for (var i = 1; i <= 8; i++) {
    sheet.autoResizeColumn(i);
  }

  // 設定數字格式
  sheet.getRange(2, 2, 資料.length, 1).setNumberFormat("#,##0");
  sheet.getRange(2, 4, 資料.length, 1).setNumberFormat("0.00");

  // 凍結標題
  sheet.setFrozenRows(1);

  Logger.log("✅ 訂單資料已建立！");
  SpreadsheetApp.getUi().alert("✅ 訂單資料已建立！\n請執行「批次計算訂單」來計算結果。");
}

// ============================================================
// 自訂選單
// ============================================================

/**
 * 開啟試算表時建立選單
 */
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu("📚 Session 2 工具")
    .addItem("📦 初始化訂單資料", "初始化訂單資料")
    .addItem("🧮 批次計算訂單", "批次計算訂單")
    .addItem("📊 訂單分析", "訂單分析")
    .addSeparator()
    .addItem("🔍 除錯示範", "除錯示範")
    .addToUi();
}
