// ============================================================
// Session 8：條件式格式化自動化
// 日期：115/05/23（六）13:30~16:30
// 講師：林冠廷
// ============================================================
// 本課程涵蓋：
//   1. 條件式格式化自動化
//   2. 執行限制（避免重複觸發）
//   3. 多層迴圈處理格式
//   4. 實作：自動格式化報表
// ============================================================

// ============================================================
// 第一部分：條件式格式化
// ============================================================

/**
 * 使用程式碼建立條件式格式規則
 * 說明：只有透過 Apps Script 才能動態新增條件式格式
 */
function 條件式格式化示範() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("業績報表");
  if (!sheet) {
    SpreadsheetApp.getUi().alert("❌ 請先執行「初始化業績資料」");
    return;
  }

  // 清除舊的條件式格式（避免重複）
  sheet.clearConditionalFormatRules();

  var 規則列表 = [];

  // --- 規則 1：業績 >= 100000 → 綠色背景 ---
  var 綠色規則 = SpreadsheetApp.newConditionalFormatRule()
    .whenNumberGreaterThanOrEqualTo(100000)
    .setBackground("#c8e6c9")        // 淺綠
    .setFontColor("#1b5e20")         // 深綠
    .setRanges([sheet.getRange("C2:C13")])  // 業績欄
    .build();
  規則列表.push(綠色規則);

  // --- 規則 2：業績 < 50000 → 紅色背景 ---
  var 紅色規則 = SpreadsheetApp.newConditionalFormatRule()
    .whenNumberLessThan(50000)
    .setBackground("#ffcdd2")        // 淺紅
    .setFontColor("#b71c1c")         // 深紅
    .setRanges([sheet.getRange("C2:C13")])
    .build();
  規則列表.push(紅色規則);

  // --- 規則 3：達成率 >= 100% → 粗體 + 綠色 ---
  var 達成規則 = SpreadsheetApp.newConditionalFormatRule()
    .whenNumberGreaterThanOrEqualTo(1)
    .setBackground("#e8f5e9")
    .setBold(true)
    .setRanges([sheet.getRange("E2:E13")])  // 達成率欄
    .build();
  規則列表.push(達成規則);

  // --- 規則 4：達成率 < 80% → 紅色警示 ---
  var 警示規則 = SpreadsheetApp.newConditionalFormatRule()
    .whenNumberLessThan(0.8)
    .setBackground("#ffebee")
    .setFontColor("#c62828")
    .setBold(true)
    .setRanges([sheet.getRange("E2:E13")])
    .build();
  規則列表.push(警示規則);

  // --- 規則 5：文字條件 — 包含「未達標」→ 紅底 ---
  var 文字規則 = SpreadsheetApp.newConditionalFormatRule()
    .whenTextContains("未達標")
    .setBackground("#ff8a80")
    .setFontColor("#ffffff")
    .setRanges([sheet.getRange("F2:F13")])  // 狀態欄
    .build();
  規則列表.push(文字規則);

  // --- 規則 6：文字條件 — 包含「達標」→ 綠底 ---
  var 達標規則 = SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo("達標")
    .setBackground("#69f0ae")
    .setRanges([sheet.getRange("F2:F13")])
    .build();
  規則列表.push(達標規則);

  // 套用所有規則
  sheet.setConditionalFormatRules(規則列表);

  Logger.log("✅ 已建立 " + 規則列表.length + " 條條件式格式規則");
  SpreadsheetApp.getUi().alert("✅ 條件式格式化已套用！\n共 " + 規則列表.length + " 條規則。");
}

// ============================================================
// 第二部分：避免重複觸發
// ============================================================

/**
 * onEdit 觸發器：儲存格編輯時自動處理
 * 說明：示範如何避免觸發器的重複執行問題
 *
 * @param {Object} e - 事件物件，包含編輯資訊
 */
function onEdit(e) {
  try {
    // 取得編輯資訊
    var sheet = e.source.getActiveSheet();
    var range = e.range;
    var 工作表名 = sheet.getName();

    // ⚠️ 技巧 1：限制觸發範圍（只在特定工作表執行）
    if (工作表名 !== "業績報表") return;

    // ⚠️ 技巧 2：限制觸發的欄位（只在 C 欄=業績時觸發）
    var 編輯欄 = range.getColumn();
    if (編輯欄 !== 3) return; // C 欄 = 第 3 欄

    // ⚠️ 技巧 3：避免標題列觸發
    var 編輯列 = range.getRow();
    if (編輯列 < 2) return;

    // 執行計算：更新達成率和狀態
    var 目標 = sheet.getRange(編輯列, 4).getValue(); // D 欄：目標
    var 業績 = e.value ? Number(e.value) : 0;

    if (目標 > 0) {
      var 達成率 = 業績 / 目標;
      sheet.getRange(編輯列, 5).setValue(達成率); // E 欄：達成率

      // 判斷狀態
      var 狀態 = 達成率 >= 1 ? "達標" : (達成率 >= 0.8 ? "接近" : "未達標");
      sheet.getRange(編輯列, 6).setValue(狀態); // F 欄：狀態
    }

  } catch (錯誤) {
    // ⚠️ 技巧 4：onEdit 中不能使用 SpreadsheetApp.getUi()
    // 只能用 Logger 記錄錯誤
    Logger.log("onEdit 錯誤：" + 錯誤.message);
  }
}

/**
 * 使用鎖定 (Lock) 防止並發執行
 * 說明：當多人同時操作可能觸發相同腳本時使用
 */
function 安全批次更新() {
  // 取得腳本鎖定
  var lock = LockService.getScriptLock();

  try {
    // 嘗試取得鎖定（等待最多 30 秒）
    lock.waitLock(30000);

    Logger.log("已取得鎖定，開始執行...");

    // === 主要邏輯 ===
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("業績報表");
    if (!sheet) return;

    var 資料 = sheet.getDataRange().getValues();
    var 達成率列 = [];
    var 狀態列 = [];

    for (var i = 1; i < 資料.length; i++) {
      var 業績 = 資料[i][2]; // C 欄
      var 目標 = 資料[i][3]; // D 欄
      var 達成率 = 目標 > 0 ? 業績 / 目標 : 0;
      var 狀態 = 達成率 >= 1 ? "達標" : (達成率 >= 0.8 ? "接近" : "未達標");

      達成率列.push([達成率]);
      狀態列.push([狀態]);
    }

    var 筆數 = 達成率列.length;
    sheet.getRange(2, 5, 筆數, 1).setValues(達成率列);
    sheet.getRange(2, 6, 筆數, 1).setValues(狀態列);

    // 數字格式
    sheet.getRange(2, 5, 筆數, 1).setNumberFormat("0.0%");

    Logger.log("✅ 安全批次更新完成");
    SpreadsheetApp.getUi().alert("✅ 業績達成率已更新！");

  } catch (錯誤) {
    Logger.log("❌ 錯誤：" + 錯誤.message);
  } finally {
    // ⚠️ 一定要釋放鎖定！
    lock.releaseLock();
  }
}

// ============================================================
// 第三部分：多層迴圈處理格式
// ============================================================

/**
 * 用多層迴圈建立熱力圖效果
 * 說明：根據數值大小自動套用漸層顏色
 */
function 建立熱力圖() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName("熱力圖");
    if (sheet) sheet.clear(); else sheet = ss.insertSheet("熱力圖");

    // 建立月度×部門資料表
    var 部門 = ["業務部", "行銷部", "研發部", "人資部", "財務部", "客服部"];
    var 月份 = ["1月", "2月", "3月", "4月", "5月", "6月",
                "7月", "8月", "9月", "10月", "11月", "12月"];

    // 標題
    sheet.getRange("A1").setValue("📊 部門月度業績熱力圖");
    sheet.getRange("A1").setFontSize(14).setFontWeight("bold");

    // 月份標題列
    sheet.getRange(3, 1).setValue("部門 \\ 月份");
    for (var m = 0; m < 月份.length; m++) {
      sheet.getRange(3, m + 2).setValue(月份[m]);
    }
    sheet.getRange(3, 1, 1, 月份.length + 1).setFontWeight("bold").setBackground("#e0e0e0");

    // 部門名稱 + 隨機數據
    var 最大值 = 0;
    var 最小值 = Infinity;
    var 數據 = [];

    for (var d = 0; d < 部門.length; d++) {
      sheet.getRange(4 + d, 1).setValue(部門[d]);
      sheet.getRange(4 + d, 1).setFontWeight("bold");

      var 列數據 = [];
      for (var m2 = 0; m2 < 月份.length; m2++) {
        var 值 = Math.floor(Math.random() * 150000) + 30000;
        sheet.getRange(4 + d, m2 + 2).setValue(值);
        列數據.push(值);

        if (值 > 最大值) 最大值 = 值;
        if (值 < 最小值) 最小值 = 值;
      }
      數據.push(列數據);
    }

    // === 用多層迴圈套用熱力圖顏色 ===
    for (var row = 0; row < 部門.length; row++) {
      for (var col = 0; col < 月份.length; col++) {
        var 值 = 數據[row][col];

        // 計算數值在範圍中的比例（0~1）
        var 比例 = (值 - 最小值) / (最大值 - 最小值);

        // 根據比例產生顏色（紅→黃→綠漸層）
        var 顏色 = 取得漸層顏色(比例);

        var cell = sheet.getRange(4 + row, col + 2);
        cell.setBackground(顏色);
        cell.setNumberFormat("#,##0");
        cell.setHorizontalAlignment("right");

        // 數值較暗時用白色字
        if (比例 < 0.3) {
          cell.setFontColor("#ffffff");
        }
      }
    }

    // 欄寬
    sheet.setColumnWidth(1, 100);
    for (var c = 2; c <= 13; c++) sheet.setColumnWidth(c, 80);

    Logger.log("✅ 熱力圖已生成！");
    SpreadsheetApp.getUi().alert("✅ 月度業績熱力圖已生成！");

  } catch (錯誤) {
    Logger.log("❌ 錯誤：" + 錯誤.message);
  }
}

/**
 * 根據 0~1 的比例回傳紅→黃→綠的漸層顏色
 * @param {number} 比例 - 0（最低，紅色）到 1（最高，綠色）
 * @returns {string} Hex 顏色碼
 */
function 取得漸層顏色(比例) {
  var r, g, b;

  if (比例 < 0.5) {
    // 紅 → 黃（0~0.5）
    r = 239;
    g = Math.round(108 + (比例 * 2) * (183 - 108));
    b = Math.round(83 + (比例 * 2) * (77 - 83));
  } else {
    // 黃 → 綠（0.5~1）
    r = Math.round(239 - ((比例 - 0.5) * 2) * (239 - 76));
    g = Math.round(183 + ((比例 - 0.5) * 2) * (175 - 183));
    b = Math.round(77 + ((比例 - 0.5) * 2) * (80 - 77));
  }

  return "#" + 轉Hex(r) + 轉Hex(g) + 轉Hex(b);
}

function 轉Hex(數值) {
  var hex = Math.max(0, Math.min(255, 數值)).toString(16);
  return hex.length === 1 ? "0" + hex : hex;
}

// ============================================================
// 第四部分：實作 — 自動格式化報表
// ============================================================

/**
 * 🔧 完整報表自動格式化
 * 說明：讀取業績資料，自動計算、套用格式、加入條件式格式
 */
function 自動格式化報表() {
  try {
    // Step 1：更新達成率
    安全批次更新();

    // Step 2：套用條件式格式
    條件式格式化示範();

    // Step 3：格式美化
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("業績報表");
    if (!sheet) return;

    // 標題列
    var 標題範圍 = sheet.getRange("A1:F1");
    標題範圍.setBackground("#283593");
    標題範圍.setFontColor("#ffffff");
    標題範圍.setFontWeight("bold");
    標題範圍.setFontSize(11);
    標題範圍.setHorizontalAlignment("center");
    sheet.setRowHeight(1, 35);

    // 數字格式
    var 最後列 = sheet.getLastRow();
    sheet.getRange(2, 3, 最後列 - 1, 1).setNumberFormat("#,##0");    // 業績
    sheet.getRange(2, 4, 最後列 - 1, 1).setNumberFormat("#,##0");    // 目標
    sheet.getRange(2, 5, 最後列 - 1, 1).setNumberFormat("0.0%");    // 達成率

    // 框線
    sheet.getDataRange().setBorder(true, true, true, true, true, true,
      "#90caf9", SpreadsheetApp.BorderStyle.SOLID);

    // 凍結標題
    sheet.setFrozenRows(1);
    for (var c = 1; c <= 6; c++) sheet.autoResizeColumn(c);

    Logger.log("✅ 報表格式化完成！");

  } catch (錯誤) {
    Logger.log("❌ 錯誤：" + 錯誤.message);
  }
}

// ============================================================
// 初始化範例資料
// ============================================================

function 初始化業績資料() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("業績報表");
  if (!sheet) sheet = ss.insertSheet("業績報表"); else sheet.clear();

  var 標題 = [["業務人員", "月份", "業績", "目標", "達成率", "狀態"]];
  var 業務 = ["王小明", "李小華", "張美玲", "陳大文"];
  var 資料 = [];

  for (var m = 1; m <= 3; m++) {
    for (var b = 0; b < 業務.length; b++) {
      var 目標 = 80000 + Math.floor(Math.random() * 40000);
      var 業績 = Math.floor(目標 * (0.5 + Math.random() * 0.8));
      var 達成率 = 業績 / 目標;
      var 狀態 = 達成率 >= 1 ? "達標" : (達成率 >= 0.8 ? "接近" : "未達標");
      資料.push([業務[b], m + "月", 業績, 目標, 達成率, 狀態]);
    }
  }

  sheet.getRange(1, 1, 1, 6).setValues(標題);
  sheet.getRange(2, 1, 資料.length, 6).setValues(資料);
  sheet.getRange("A1:F1").setBackground("#283593").setFontColor("#fff").setFontWeight("bold");
  sheet.getRange(2, 3, 資料.length, 2).setNumberFormat("#,##0");
  sheet.getRange(2, 5, 資料.length, 1).setNumberFormat("0.0%");
  sheet.setFrozenRows(1);
  for (var c = 1; c <= 6; c++) sheet.autoResizeColumn(c);

  SpreadsheetApp.getUi().alert("✅ 業績資料已建立！共 " + 資料.length + " 筆");
}

// ============================================================
// 自訂選單
// ============================================================

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu("📚 Session 8 工具")
    .addItem("📦 初始化業績資料", "初始化業績資料")
    .addItem("🎨 條件式格式化", "條件式格式化示範")
    .addItem("📊 安全批次更新", "安全批次更新")
    .addSeparator()
    .addItem("🌡️ 建立熱力圖", "建立熱力圖")
    .addItem("📋 自動格式化報表", "自動格式化報表")
    .addToUi();
}
