// ============================================================
// Session 7：自動格式設定
// 日期：115/05/23（六）09:00~12:00
// 講師：林冠廷
// ============================================================
// 本課程涵蓋：
//   1. Apps Script 介面進階技巧
//   2. 文字、數字、日期自動套用格式
//   3. 格式設定方法大全
//   4. 實作：自動生成標題格式表格
// ============================================================

// ============================================================
// 第一部分：文字格式設定
// ============================================================

/**
 * 文字格式完整示範
 */
function 文字格式設定() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("格式範例");
  if (!sheet) {
    SpreadsheetApp.getUi().alert("❌ 請先執行「初始化格式資料」");
    return;
  }

  // --- 字體設定 ---
  var A1 = sheet.getRange("A1");
  A1.setValue("字體示範");
  A1.setFontFamily("微軟正黑體");    // 字體
  A1.setFontSize(14);                 // 字體大小
  A1.setFontWeight("bold");           // 粗體
  A1.setFontStyle("italic");          // 斜體
  A1.setFontColor("#1a73e8");         // 字體顏色

  // --- 底線與刪除線 ---
  var A2 = sheet.getRange("A2");
  A2.setValue("底線文字");
  A2.setFontLine("underline");        // 底線

  var A3 = sheet.getRange("A3");
  A3.setValue("刪除線文字");
  A3.setFontLine("line-through");     // 刪除線

  // --- 文字對齊 ---
  sheet.getRange("B1").setValue("靠左").setHorizontalAlignment("left");
  sheet.getRange("B2").setValue("置中").setHorizontalAlignment("center");
  sheet.getRange("B3").setValue("靠右").setHorizontalAlignment("right");

  // --- 垂直對齊 ---
  sheet.setRowHeights(1, 3, 50);
  sheet.getRange("C1").setValue("上").setVerticalAlignment("top");
  sheet.getRange("C2").setValue("中").setVerticalAlignment("middle");
  sheet.getRange("C3").setValue("下").setVerticalAlignment("bottom");

  // --- 自動換行 ---
  var D1 = sheet.getRange("D1");
  D1.setValue("這是一段很長的文字，設定自動換行後會在儲存格內換行顯示");
  D1.setWrap(true);                   // 開啟自動換行
  sheet.setColumnWidth(4, 200);

  Logger.log("✅ 文字格式設定完成");
}

// ============================================================
// 第二部分：數字格式設定
// ============================================================

/**
 * 數字格式完整示範
 */
function 數字格式設定() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("格式範例");
  if (!sheet) return;

  // 在 F 欄示範各種數字格式
  var 起始列 = 1;
  var 數值 = 1234567.89;

  var 格式範例 = [
    ["數字格式", "結果"],
    ["原始值", 數值],
    ["千分位 (#,##0)", 數值],
    ["小數兩位 (#,##0.00)", 數值],
    ["百分比 (0.00%)", 0.8567],
    ["貨幣 ($#,##0)", 數值],
    ["NT$ ($#,##0)", 數值],
    ["日期 (yyyy/mm/dd)", new Date()],
    ["時間 (hh:mm:ss)", new Date()],
    ["日期時間", new Date()],
    ["自訂 (第 0 名)", 42]
  ];

  // 寫入範例
  sheet.getRange("F1:G1").setValues([格式範例[0]]);
  sheet.getRange("F1:G1").setBackground("#1a73e8").setFontColor("#fff").setFontWeight("bold");

  for (var i = 1; i < 格式範例.length; i++) {
    sheet.getRange(i + 1, 6).setValue(格式範例[i][0]);
    sheet.getRange(i + 1, 7).setValue(格式範例[i][1]);
  }

  // 套用不同的數字格式
  sheet.getRange("G3").setNumberFormat("#,##0");
  sheet.getRange("G4").setNumberFormat("#,##0.00");
  sheet.getRange("G5").setNumberFormat("0.00%");
  sheet.getRange("G6").setNumberFormat("$#,##0");
  sheet.getRange("G7").setNumberFormat("NT$#,##0");
  sheet.getRange("G8").setNumberFormat("yyyy/mm/dd");
  sheet.getRange("G9").setNumberFormat("hh:mm:ss");
  sheet.getRange("G10").setNumberFormat("yyyy/mm/dd hh:mm");
  sheet.getRange("G11").setNumberFormat('"第 "0" 名"');

  // 自動調整欄寬
  sheet.autoResizeColumn(6);
  sheet.autoResizeColumn(7);

  Logger.log("✅ 數字格式設定完成");
}

// ============================================================
// 第三部分：背景與框線設定
// ============================================================

/**
 * 背景與框線示範
 */
function 背景框線設定() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("格式範例");
  if (!sheet) return;

  // --- 背景色 ---
  sheet.getRange("A15").setValue("背景色示範");
  sheet.getRange("A15").setFontWeight("bold").setFontSize(12);

  var 顏色 = ["#e8f5e9", "#e3f2fd", "#fce4ec", "#fff3e0", "#f3e5f5", "#e0f7fa"];
  var 色名 = ["綠色系", "藍色系", "粉紅系", "橘色系", "紫色系", "青色系"];

  for (var i = 0; i < 顏色.length; i++) {
    sheet.getRange(16, i + 1).setValue(色名[i]);
    sheet.getRange(16, i + 1).setBackground(顏色[i]);
    sheet.getRange(16, i + 1).setHorizontalAlignment("center");
  }

  // --- 框線樣式 ---
  sheet.getRange("A18").setValue("框線樣式示範");
  sheet.getRange("A18").setFontWeight("bold").setFontSize(12);

  // 實線框線
  var r1 = sheet.getRange("A19:C21");
  r1.setValue("實線框線");
  r1.setBorder(true, true, true, true, true, true,
    "#000000", SpreadsheetApp.BorderStyle.SOLID);

  // 粗線框線
  var r2 = sheet.getRange("A23:C25");
  r2.setValue("粗線框線");
  r2.setBorder(true, true, true, true, true, true,
    "#1a73e8", SpreadsheetApp.BorderStyle.SOLID_MEDIUM);

  // 虛線框線
  var r3 = sheet.getRange("A27:C29");
  r3.setValue("虛線框線");
  r3.setBorder(true, true, true, true, true, true,
    "#34a853", SpreadsheetApp.BorderStyle.DASHED);

  Logger.log("✅ 背景框線設定完成");
}

// ============================================================
// 第四部分：實作 — 自動生成標題格式表格
// ============================================================

/**
 * 🔧 一鍵生成專業格式表格
 * 說明：輸入資料後，自動套用專業的表格格式
 */
function 自動生成格式表格() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();

    // 建立新的「格式化報表」工作表
    var 表名 = "格式化報表";
    var sheet = ss.getSheetByName(表名);
    if (sheet) sheet.clear(); else sheet = ss.insertSheet(表名);

    // === 表格標題 ===
    sheet.getRange("A1:F1").merge();
    sheet.getRange("A1").setValue("📊 2026年度部門預算報表");
    sheet.getRange("A1").setFontSize(18).setFontWeight("bold")
         .setHorizontalAlignment("center").setFontColor("#1a237e");
    sheet.setRowHeight(1, 50);

    // 副標題
    sheet.getRange("A2:F2").merge();
    sheet.getRange("A2").setValue("報告日期：" +
      Utilities.formatDate(new Date(), "Asia/Taipei", "yyyy年MM月dd日"));
    sheet.getRange("A2").setHorizontalAlignment("center")
         .setFontColor("#666666").setFontSize(10);

    // === 欄位標題列 ===
    var 標題 = [["部門", "Q1預算", "Q2預算", "Q3預算", "Q4預算", "年度合計"]];
    sheet.getRange("A4:F4").setValues(標題);
    sheet.getRange("A4:F4")
      .setBackground("#1a237e")
      .setFontColor("#000000")
      .setFontWeight("bold")
      .setFontSize(11)
      .setHorizontalAlignment("center");
    sheet.setRowHeight(4, 35);

    // === 資料 ===
    var 資料 = [
      ["業務部", 1500000, 1800000, 2000000, 2200000],
      ["行銷部", 800000, 900000, 1000000, 1100000],
      ["研發部", 2500000, 2600000, 2700000, 2800000],
      ["人資部", 500000, 520000, 540000, 560000],
      ["財務部", 400000, 420000, 430000, 450000],
      ["客服部", 600000, 650000, 700000, 750000]
    ];

    // 計算年度合計並加入
    var 完整資料 = 資料.map(function(row) {
      var 合計 = row[1] + row[2] + row[3] + row[4];
      return [row[0], row[1], row[2], row[3], row[4], 合計];
    });

    sheet.getRange(5, 1, 完整資料.length, 6).setValues(完整資料);

    // === 資料列格式 ===
    var 資料範圍 = sheet.getRange(5, 1, 完整資料.length, 6);

    // 交替行背景色（斑馬紋）
    for (var i = 0; i < 完整資料.length; i++) {
      var 行範圍 = sheet.getRange(5 + i, 1, 1, 6);
      if (i % 2 === 0) {
        行範圍.setBackground("#f5f5f5");
      } else {
        行範圍.setBackground("#ffffff");
      }
    }

    // 數字格式：千分位加 NT$
    sheet.getRange(5, 2, 完整資料.length, 5).setNumberFormat("NT$#,##0");
    sheet.getRange(5, 2, 完整資料.length, 5).setHorizontalAlignment("right");

    // 年度合計列特殊標記
    sheet.getRange(5, 6, 完整資料.length, 1).setFontWeight("bold");

    // === 合計列 ===
    var 合計列 = 5 + 完整資料.length;
    sheet.getRange(合計列, 1).setValue("總 計");
    sheet.getRange(合計列, 1).setFontWeight("bold").setHorizontalAlignment("center");

    for (var col = 2; col <= 6; col++) {
      var 欄字母 = String.fromCharCode(64 + col); // B, C, D, E, F
      var 公式 = "=SUM(" + 欄字母 + "5:" + 欄字母 + (合計列 - 1) + ")";
      sheet.getRange(合計列, col).setFormula(公式);
    }

    sheet.getRange(合計列, 1, 1, 6)
      .setBackground("#e8eaf6")
      .setFontWeight("bold")
      .setBorder(true, false, true, false, false, false,
        "#1a237e", SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
    sheet.getRange(合計列, 2, 1, 5).setNumberFormat("NT$#,##0");

    // === 整體框線 ===
    sheet.getRange(4, 1, 完整資料.length + 2, 6)
      .setBorder(true, true, true, true, true, true,
        "#bdbdbd", SpreadsheetApp.BorderStyle.SOLID);

    // === 欄寬設定 ===
    sheet.setColumnWidth(1, 100);
    for (var c = 2; c <= 6; c++) {
      sheet.setColumnWidth(c, 130);
    }

    // 凍結標題
    sheet.setFrozenRows(4);

    Logger.log("✅ 格式化報表已生成！");
    SpreadsheetApp.getUi().alert("✅ 專業格式報表已生成！\n請查看「格式化報表」工作表。");

  } catch (錯誤) {
    Logger.log("❌ 錯誤：" + 錯誤.message);
    SpreadsheetApp.getUi().alert("❌ 錯誤：" + 錯誤.message);
  }
}

/**
 * 一鍵美化當前工作表
 * 說明：選取任何含資料的工作表，自動套用專業格式
 */
function 一鍵美化() {
  try {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    var 資料範圍 = sheet.getDataRange();
    var 最後列 = sheet.getLastRow();
    var 最後欄 = sheet.getLastColumn();

    if (最後列 < 2 || 最後欄 < 1) {
      SpreadsheetApp.getUi().alert("⚠️ 工作表中沒有足夠的資料");
      return;
    }

    // 標題列格式
    var 標題 = sheet.getRange(1, 1, 1, 最後欄);
    標題.setBackground("#2196f3");
    標題.setFontColor("#000000");
    標題.setFontWeight("bold");
    標題.setFontSize(11);
    標題.setHorizontalAlignment("center");
    sheet.setRowHeight(1, 35);

    // 斑馬紋
    for (var i = 2; i <= 最後列; i++) {
      var 行 = sheet.getRange(i, 1, 1, 最後欄);
      行.setBackground(i % 2 === 0 ? "#e3f2fd" : "#ffffff");
    }

    // 框線
    資料範圍.setBorder(true, true, true, true, true, true,
      "#90caf9", SpreadsheetApp.BorderStyle.SOLID);

    // 自動調整欄寬
    for (var c = 1; c <= 最後欄; c++) {
      sheet.autoResizeColumn(c);
    }

    // 凍結標題列
    sheet.setFrozenRows(1);

    SpreadsheetApp.getUi().alert("✅ 工作表已美化完成！");

  } catch (錯誤) {
    Logger.log("❌ 錯誤：" + 錯誤.message);
  }
}

// ============================================================
// 初始化範例資料
// ============================================================

function 初始化格式資料() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("格式範例");
  if (!sheet) sheet = ss.insertSheet("格式範例"); else sheet.clear();

  // 基本提示
  sheet.getRange("A1").setValue("請執行各個格式設定函數來查看效果");
  sheet.getRange("A1").setFontSize(12).setFontWeight("bold");

  SpreadsheetApp.getUi().alert("✅ 格式範例工作表已建立！\n請依序執行各格式設定函數。");
}

// ============================================================
// 自訂選單
// ============================================================

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu("📚 Session 7 工具")
    .addItem("📦 初始化格式範例", "初始化格式資料")
    .addSeparator()
    .addItem("✏️ 文字格式設定", "文字格式設定")
    .addItem("🔢 數字格式設定", "數字格式設定")
    .addItem("🎨 背景框線設定", "背景框線設定")
    .addSeparator()
    .addItem("📊 自動生成格式表格", "自動生成格式表格")
    .addItem("✨ 一鍵美化當前表", "一鍵美化")
    .addToUi();
}
