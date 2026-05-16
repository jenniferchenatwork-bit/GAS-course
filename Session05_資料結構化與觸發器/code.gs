// ============================================================
// Session 5：資料結構化與觸發器
// 日期：115/05/16（六）09:00~12:00
// 講師：林冠廷
// ============================================================
// 本課程涵蓋：
//   1. 基本語法延伸（function & return）
//   2. 資料讀取與結構化 (Array)
//   3. 觸發器應用（定期更新資料）
//   4. 實作：定時抓取表格並生成清單
// ============================================================

// ============================================================
// 第一部分：函數進階 — function & return
// ============================================================

/**
 * 多回傳值的函數設計
 * 說明：用物件 (Object) 或陣列回傳多個值
 */
function 函數進階示範() {
  // --- 回傳物件 ---
  var 員工資訊 = 取得員工統計();
  Logger.log("總人數：" + 員工資訊.總人數);
  Logger.log("平均薪資：" + 員工資訊.平均薪資);
  Logger.log("最高薪資：" + 員工資訊.最高薪資);

  // --- 函數作為參數（回呼函數）---
  var 數字 = [3, 1, 4, 1, 5, 9, 2, 6];
  var 篩選結果 = 篩選陣列(數字, function(x) {
    return x > 3;  // 只保留大於 3 的數字
  });
  Logger.log("大於 3 的數字：" + 篩選結果); // [4, 5, 9, 6]
}

/**
 * 從試算表讀取員工資料並計算統計值
 * @returns {Object} 包含統計資訊的物件
 */
function 取得員工統計() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("專案人員");
  if (!sheet) {
    return { 總人數: 0, 平均薪資: 0, 最高薪資: 0 };
  }

  var 資料 = sheet.getDataRange().getValues();
  var 薪資列表 = [];

  for (var i = 1; i < 資料.length; i++) {
    薪資列表.push(資料[i][3]); // D 欄：月薪
  }

  // 計算統計值
  var 總人數 = 薪資列表.length;
  var 薪資總和 = 薪資列表.reduce(function(sum, val) { return sum + val; }, 0);
  var 平均薪資 = Math.round(薪資總和 / 總人數);
  var 最高薪資 = Math.max.apply(null, 薪資列表);
  var 最低薪資 = Math.min.apply(null, 薪資列表);

  // 回傳一個物件（包含多個值）
  return {
    總人數: 總人數,
    平均薪資: 平均薪資,
    最高薪資: 最高薪資,
    最低薪資: 最低薪資,
    薪資總和: 薪資總和
  };
}

/**
 * 篩選陣列的通用函數
 * @param {Array} 陣列 - 要篩選的陣列
 * @param {Function} 條件函數 - 篩選條件（回傳 true/false）
 * @returns {Array} 篩選後的陣列
 */
function 篩選陣列(陣列, 條件函數) {
  var 結果 = [];
  for (var i = 0; i < 陣列.length; i++) {
    if (條件函數(陣列[i])) {
      結果.push(陣列[i]);
    }
  }
  return 結果;
}

// ============================================================
// 第二部分：資料結構化
// ============================================================

/**
 * 將試算表資料轉換為結構化物件陣列
 * 說明：比二維陣列更直覺、更好維護
 */
function 資料結構化示範() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("專案人員");
  if (!sheet) {
    SpreadsheetApp.getUi().alert("❌ 請先執行「初始化專案資料」");
    return;
  }

  var 資料 = sheet.getDataRange().getValues();
  var 標題列 = 資料[0]; // ["姓名", "部門", "職稱", "月薪", ...]

  // 將二維陣列轉為物件陣列
  var 員工列表 = [];
  for (var i = 1; i < 資料.length; i++) {
    var 員工 = {};
    for (var j = 0; j < 標題列.length; j++) {
      員工[標題列[j]] = 資料[i][j];
    }
    員工列表.push(員工);
  }

  // 現在可以用 .屬性名 來存取，更直覺！
  Logger.log("===== 結構化資料 =====");
  for (var k = 0; k < 員工列表.length; k++) {
    Logger.log(員工列表[k]["姓名"] + " - " + 員工列表[k]["部門"] + " - NT$" + 員工列表[k]["月薪"]);
  }

  // 進階：依部門分組
  var 部門分組 = 依欄位分組(員工列表, "部門");

  Logger.log("\n===== 部門分組 =====");
  for (var 部門 in 部門分組) {
    Logger.log(部門 + "：" + 部門分組[部門].length + " 人");
    for (var m = 0; m < 部門分組[部門].length; m++) {
      Logger.log("  - " + 部門分組[部門][m]["姓名"]);
    }
  }
}

/**
 * 依指定欄位將資料分組
 * @param {Array<Object>} 資料 - 物件陣列
 * @param {string} 欄位名 - 要分組的欄位名稱
 * @returns {Object} 分組結果
 */
function 依欄位分組(資料, 欄位名) {
  var 分組 = {};
  for (var i = 0; i < 資料.length; i++) {
    var 值 = 資料[i][欄位名];
    if (!分組[值]) {
      分組[值] = [];
    }
    分組[值].push(資料[i]);
  }
  return 分組;
}

// ============================================================
// 第三部分：觸發器應用 — 定期更新資料
// ============================================================

/**
 * 設定每小時自動更新的觸發器
 */
function 設定定期更新觸發器() {
  // 清除同名舊觸發器
  var triggers = ScriptApp.getProjectTriggers();
  for (var i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === "定時更新專案狀態") {
      ScriptApp.deleteTrigger(triggers[i]);
    }
  }

  // 建立每小時觸發器
  ScriptApp.newTrigger("定時更新專案狀態")
    .timeBased()
    .everyHours(1)
    .create();

  Logger.log("✅ 每小時更新觸發器已設定");
  SpreadsheetApp.getUi().alert("✅ 已設定每小時自動更新！");
}

/**
 * 定時更新專案狀態
 * 說明：自動計算專案進度、逾期檢查，並生成摘要
 */
function 定時更新專案狀態() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var 專案表 = ss.getSheetByName("專案追蹤");
    if (!專案表) return;

    var 資料 = 專案表.getDataRange().getValues();
    var 今天 = new Date();
    var 更新結果 = [];

    for (var i = 1; i < 資料.length; i++) {
      var 截止日 = new Date(資料[i][4]); // E 欄：截止日期
      var 進度 = 資料[i][5];             // F 欄：完成進度
      var 狀態;

      if (進度 >= 100) {
        狀態 = "✅ 已完成";
      } else if (今天 > 截止日) {
        狀態 = "🔴 已逾期";
      } else {
        var 剩餘天數 = Math.ceil((截止日 - 今天) / (1000 * 60 * 60 * 24));
        if (剩餘天數 <= 3) {
          狀態 = "🟡 即將到期 (" + 剩餘天數 + "天)";
        } else {
          狀態 = "🟢 進行中 (" + 剩餘天數 + "天)";
        }
      }

      更新結果.push([狀態]);
    }

    // 批次寫入狀態
    if (更新結果.length > 0) {
      專案表.getRange(2, 7, 更新結果.length, 1).setValues(更新結果);
    }

    // 更新時間戳記
    專案表.getRange("I1").setValue("最後更新");
    專案表.getRange("I2").setValue(Utilities.formatDate(今天, "Asia/Taipei", "yyyy/MM/dd HH:mm"));

    Logger.log("✅ 專案狀態已更新（" + 更新結果.length + " 筆）");

  } catch (錯誤) {
    Logger.log("❌ 定時更新錯誤：" + 錯誤.message);
  }
}

// ============================================================
// 第四部分：實作 — 定時抓取表格並生成清單
// ============================================================

/**
 * 生成部門人員清單
 * 說明：從專案人員表讀取資料，自動產生按部門分類的人員清單
 */
function 生成部門清單() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var 人員表 = ss.getSheetByName("專案人員");
    if (!人員表) {
      SpreadsheetApp.getUi().alert("❌ 請先執行「初始化專案資料」");
      return;
    }

    // 讀取並結構化資料
    var 資料 = 人員表.getDataRange().getValues();
    var 標題 = 資料[0];
    var 員工 = [];
    for (var i = 1; i < 資料.length; i++) {
      var obj = {};
      for (var j = 0; j < 標題.length; j++) {
        obj[標題[j]] = 資料[i][j];
      }
      員工.push(obj);
    }

    // 依部門分組
    var 部門 = 依欄位分組(員工, "部門");

    // 建立或清除「部門清單」工作表
    var 清單表 = ss.getSheetByName("部門清單");
    if (清單表) {
      清單表.clear();
    } else {
      清單表 = ss.insertSheet("部門清單");
    }

    // 寫入標題
    清單表.getRange("A1").setValue("📋 部門人員清單");
    清單表.getRange("A1").setFontSize(16).setFontWeight("bold");
    清單表.getRange("A2").setValue("更新時間：" + Utilities.formatDate(new Date(), "Asia/Taipei", "yyyy/MM/dd HH:mm"));

    var 目前列 = 4;

    // 遍歷每個部門
    for (var 部門名 in 部門) {
      var 成員 = 部門[部門名];

      // 部門標題
      清單表.getRange(目前列, 1).setValue("🏢 " + 部門名 + " (" + 成員.length + " 人)");
      清單表.getRange(目前列, 1, 1, 4).merge();
      清單表.getRange(目前列, 1).setFontSize(13).setFontWeight("bold").setBackground("#e3f2fd");
      目前列++;

      // 欄位標題
      清單表.getRange(目前列, 1, 1, 4).setValues([["姓名", "職稱", "月薪", "Email"]]);
      清單表.getRange(目前列, 1, 1, 4).setFontWeight("bold").setBackground("#bbdefb");
      目前列++;

      // 成員資料
      for (var n = 0; n < 成員.length; n++) {
        清單表.getRange(目前列, 1, 1, 4).setValues([
          [成員[n]["姓名"], 成員[n]["職稱"], 成員[n]["月薪"], 成員[n]["Email"]]
        ]);
        清單表.getRange(目前列, 3).setNumberFormat("#,##0");
        目前列++;
      }

      目前列++; // 空一列
    }

    // 調整欄寬
    for (var c = 1; c <= 4; c++) {
      清單表.autoResizeColumn(c);
    }

    Logger.log("✅ 部門清單已生成！");
    SpreadsheetApp.getUi().alert("✅ 部門清單已生成！請查看「部門清單」工作表。");

  } catch (錯誤) {
    Logger.log("❌ 錯誤：" + 錯誤.message);
  }
}

/**
 * 生成專案進度摘要
 */
function 生成專案摘要() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var 專案表 = ss.getSheetByName("專案追蹤");
    if (!專案表) {
      SpreadsheetApp.getUi().alert("❌ 找不到「專案追蹤」工作表");
      return;
    }

    // 先更新狀態
    定時更新專案狀態();

    var 資料 = 專案表.getDataRange().getValues();

    // 統計各狀態數量
    var 統計 = { 已完成: 0, 已逾期: 0, 即將到期: 0, 進行中: 0 };
    for (var i = 1; i < 資料.length; i++) {
      var 狀態 = String(資料[i][6]); // G 欄
      if (狀態.indexOf("已完成") >= 0) 統計.已完成++;
      else if (狀態.indexOf("已逾期") >= 0) 統計.已逾期++;
      else if (狀態.indexOf("即將到期") >= 0) 統計.即將到期++;
      else 統計.進行中++;
    }

    var 總專案 = 資料.length - 1;
    var 摘要 = "📊 專案進度摘要\n\n" +
      "總專案數：" + 總專案 + "\n" +
      "✅ 已完成：" + 統計.已完成 + "\n" +
      "🟢 進行中：" + 統計.進行中 + "\n" +
      "🟡 即將到期：" + 統計.即將到期 + "\n" +
      "🔴 已逾期：" + 統計.已逾期;

    SpreadsheetApp.getUi().alert(摘要);
    Logger.log(摘要);

  } catch (錯誤) {
    Logger.log("❌ 錯誤：" + 錯誤.message);
  }
}

/**
 * 將專案依優先級分組並生成摘要 (輸出至獨立工作表)
 */
function 生成優先級摘要() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var 專案表 = ss.getSheetByName("專案追蹤");
    if (!專案表) {
      SpreadsheetApp.getUi().alert("❌ 找不到「專案追蹤」工作表");
      return;
    }

    // 先更新狀態以確保狀態欄位有值
    定時更新專案狀態();

    var 資料 = 專案表.getDataRange().getValues();
    var 標題 = 資料[0];
    var 專案列表 = [];
    for (var i = 1; i < 資料.length; i++) {
      var obj = {};
      for (var j = 0; j < 標題.length; j++) {
        obj[標題[j]] = 資料[i][j];
      }
      專案列表.push(obj);
    }

    var 分組結果 = 依欄位分組(專案列表, "優先級");
    
    // 建立或清除「優先級摘要」工作表
    var 摘要表 = ss.getSheetByName("優先級摘要");
    if (摘要表) {
      摘要表.clear();
    } else {
      摘要表 = ss.insertSheet("優先級摘要");
    }

    摘要表.getRange("A1").setValue("📌 專案優先級摘要");
    摘要表.getRange("A1").setFontSize(16).setFontWeight("bold");
    摘要表.getRange("A2").setValue("更新時間：" + Utilities.formatDate(new Date(), "Asia/Taipei", "yyyy/MM/dd HH:mm"));

    var 目前列 = 4;
    var 順序 = ["高", "中", "低"];
    
    for (var k = 0; k < 順序.length; k++) {
      var 級別 = 順序[k];
      var 該級別專案 = 分組結果[級別] || [];
      
      摘要表.getRange(目前列, 1).setValue("🎯 優先級：" + 級別 + " (" + 該級別專案.length + " 個專案)");
      摘要表.getRange(目前列, 1, 1, 5).merge();
      摘要表.getRange(目前列, 1).setFontSize(13).setFontWeight("bold").setBackground("#ffcc80");
      目前列++;

      if (該級別專案.length > 0) {
        摘要表.getRange(目前列, 1, 1, 5).setValues([["專案名稱", "負責人", "截止日期", "進度(%)", "狀態"]]);
        摘要表.getRange(目前列, 1, 1, 5).setFontWeight("bold").setBackground("#ffe082");
        目前列++;
        
        for (var m = 0; m < 該級別專案.length; m++) {
          var 專案 = 該級別專案[m];
          var 截止日 = "";
          if (專案["截止日期"]) {
            截止日 = Utilities.formatDate(new Date(專案["截止日期"]), "Asia/Taipei", "yyyy/MM/dd");
          }
          摘要表.getRange(目前列, 1, 1, 5).setValues([
            [專案["專案名稱"], 專案["負責人"], 截止日, 專案["進度(%)"], 專案["狀態"]]
          ]);
          目前列++;
        }
      }
      目前列++; // 空一列
    }

    for (var c = 1; c <= 5; c++) 摘要表.autoResizeColumn(c);

    SpreadsheetApp.getUi().alert("✅ 優先級摘要已生成！請查看「優先級摘要」工作表。");

  } catch (錯誤) {
    Logger.log("❌ 錯誤：" + 錯誤.message);
  }
}

/**
 * 自動計算每個部門的平均薪資並產生薪資報表
 */
function 生成薪資報表() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var 人員表 = ss.getSheetByName("專案人員");
    if (!人員表) {
      SpreadsheetApp.getUi().alert("❌ 找不到「專案人員」工作表");
      return;
    }

    var 資料 = 人員表.getDataRange().getValues();
    var 標題 = 資料[0];
    var 員工 = [];
    for (var i = 1; i < 資料.length; i++) {
      var obj = {};
      for (var j = 0; j < 標題.length; j++) {
        obj[標題[j]] = 資料[i][j];
      }
      員工.push(obj);
    }

    var 部門 = 依欄位分組(員工, "部門");
    var 薪資統計 = [];
    
    for (var 部門名 in 部門) {
      var 成員 = 部門[部門名];
      var 薪資總和 = 0;
      for (var n = 0; n < 成員.length; n++) {
        薪資總和 += Number(成員[n]["月薪"]) || 0;
      }
      var 平均薪資 = Math.round(薪資總和 / 成員.length);
      薪資統計.push([部門名, 成員.length, 平均薪資]);
    }

    var 報表表 = ss.getSheetByName("薪資報表");
    if (報表表) {
      報表表.clear();
    } else {
      報表表 = ss.insertSheet("薪資報表");
    }

    報表表.getRange("A1").setValue("💰 部門平均薪資報表");
    報表表.getRange("A1").setFontSize(16).setFontWeight("bold");
    報表表.getRange("A2").setValue("更新時間：" + Utilities.formatDate(new Date(), "Asia/Taipei", "yyyy/MM/dd HH:mm"));

    var 標題列 = [["部門", "人數", "平均月薪"]];
    報表表.getRange(4, 1, 1, 3).setValues(標題列)
      .setFontWeight("bold")
      .setBackground("#c8e6c9");

    if (薪資統計.length > 0) {
      報表表.getRange(5, 1, 薪資統計.length, 3).setValues(薪資統計);
      報表表.getRange(5, 3, 薪資統計.length, 1).setNumberFormat("#,##0");
    }

    for (var c = 1; c <= 3; c++) 報表表.autoResizeColumn(c);

    SpreadsheetApp.getUi().alert("✅ 薪資報表已生成！請查看「薪資報表」工作表。");
    
  } catch (錯誤) {
    Logger.log("❌ 錯誤：" + 錯誤.message);
  }
}

// ============================================================
// 初始化範例資料
// ============================================================

function 初始化專案資料() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  // --- 專案人員 ---
  var 人員表 = ss.getSheetByName("專案人員");
  if (!人員表) 人員表 = ss.insertSheet("專案人員"); else 人員表.clear();

  var 人員標題 = [["姓名", "部門", "職稱", "月薪", "Email"]];
  var 人員資料 = [
    ["王小明", "開發部", "前端工程師", 48000, "wang@company.com"],
    ["李小華", "開發部", "後端工程師", 52000, "lee@company.com"],
    ["張美玲", "設計部", "UI設計師", 45000, "chang@company.com"],
    ["陳大文", "開發部", "全端工程師", 58000, "chen@company.com"],
    ["林小芬", "企劃部", "專案經理", 55000, "lin@company.com"],
    ["黃志偉", "設計部", "UX研究員", 46000, "huang@company.com"],
    ["劉家豪", "測試部", "QA工程師", 42000, "liu@company.com"],
    ["吳雅琪", "企劃部", "產品經理", 60000, "wu@company.com"],
    ["周建國", "測試部", "自動化測試", 44000, "chou@company.com"],
    ["許文馨", "開發部", "iOS工程師", 56000, "hsu@company.com"]
  ];

  人員表.getRange(1, 1, 1, 5).setValues(人員標題);
  人員表.getRange(2, 1, 人員資料.length, 5).setValues(人員資料);
  人員表.getRange("A1:E1").setBackground("#673ab7").setFontColor("#fff").setFontWeight("bold");
  人員表.getRange("D2:D11").setNumberFormat("#,##0");
  人員表.setFrozenRows(1);
  for (var i = 1; i <= 5; i++) 人員表.autoResizeColumn(i);

  // --- 專案追蹤 ---
  var 專案表 = ss.getSheetByName("專案追蹤");
  if (!專案表) 專案表 = ss.insertSheet("專案追蹤"); else 專案表.clear();

  var 專案標題 = [["專案名稱", "負責人", "優先級", "開始日期", "截止日期", "進度(%)", "狀態"]];
  var 專案資料 = [
    ["官網改版", "林小芬", "高", new Date(2026, 3, 1), new Date(2026, 4, 15), 80, ""],
    ["App 2.0", "吳雅琪", "高", new Date(2026, 3, 10), new Date(2026, 5, 30), 35, ""],
    ["報表系統", "陳大文", "中", new Date(2026, 4, 1), new Date(2026, 4, 20), 100, ""],
    ["客戶管理", "王小明", "中", new Date(2026, 2, 15), new Date(2026, 4, 10), 60, ""],
    ["API 串接", "李小華", "高", new Date(2026, 4, 5), new Date(2026, 4, 25), 20, ""],
    ["設計系統", "張美玲", "低", new Date(2026, 3, 20), new Date(2026, 6, 1), 45, ""],
    ["自動化測試", "劉家豪", "中", new Date(2026, 4, 1), new Date(2026, 5, 15), 15, ""],
    ["用戶研究", "黃志偉", "低", new Date(2026, 3, 1), new Date(2026, 3, 30), 100, ""]
  ];

  專案表.getRange(1, 1, 1, 7).setValues(專案標題);
  專案表.getRange(2, 1, 專案資料.length, 7).setValues(專案資料);
  專案表.getRange("A1:G1").setBackground("#e65100").setFontColor("#fff").setFontWeight("bold");
  專案表.getRange("D2:E9").setNumberFormat("yyyy/mm/dd");
  專案表.setFrozenRows(1);
  for (var j = 1; j <= 7; j++) 專案表.autoResizeColumn(j);

  SpreadsheetApp.getUi().alert("✅ 專案資料已建立！（專案人員 + 專案追蹤）");
}

// ============================================================
// 自訂選單
// ============================================================

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu("📚 Session 5 工具")
    .addItem("📦 初始化專案資料", "初始化專案資料")
    .addItem("📋 生成部門清單", "生成部門清單")
    .addItem("📊 更新專案狀態", "定時更新專案狀態")
    .addItem("📈 專案進度摘要", "生成專案摘要")
    .addItem("📌 專案優先級摘要", "生成優先級摘要")
    .addItem("💰 產生薪資報表", "生成薪資報表")
    .addSeparator()
    .addItem("🔬 函數進階示範", "函數進階示範")
    .addItem("🗂️ 資料結構化示範", "資料結構化示範")
    .addItem("⏰ 設定定期更新", "設定定期更新觸發器")
    .addToUi();
}
